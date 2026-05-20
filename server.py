"""
NQ Auto Journal – Webhook Server
Receives TradingView alerts and writes trade data into Excel automatically.
"""

from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os, logging
from datetime import datetime
from zoneinfo import ZoneInfo

# ═══════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════
EXCEL_PATH    = os.getenv("EXCEL_PATH",      "NQ_Trading_Journal.xlsx")
PORT          = int(os.getenv("PORT",        5000))
SECRET_TOKEN  = os.getenv("WEBHOOK_SECRET",  "")      # Optional security token
TIMEZONE      = ZoneInfo("America/New_York")           # Adjust if needed
LOG_FILE      = "journal.log"

# ═══════════════════════════════════════════════════════════════
#  LOGGING
# ═══════════════════════════════════════════════════════════════
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler()]
)
log = logging.getLogger(__name__)

app = Flask(__name__)

# ═══════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════

def get_session(dt: datetime) -> str:
    """Detect trading session from UTC datetime"""
    h = dt.astimezone(ZoneInfo("UTC")).hour
    if  1 <= h <  8: return "Asia"
    if  8 <= h < 13: return "London"
    if 13 <= h < 15: return "NY Open"
    if 15 <= h < 21: return "New York"
    return "After Hours"

def parse_time(raw) -> datetime:
    """Parse TradingView timestamp (ms or ISO string)"""
    try:
        return datetime.fromtimestamp(float(raw) / 1000, tz=ZoneInfo("UTC"))
    except (ValueError, TypeError):
        pass
    try:
        return datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
    except Exception:
        return datetime.now(ZoneInfo("UTC"))

# Minimal cell styling
C_MID  = "161B22"
C_LITE = "21262D"
C_WH   = "E6EDF3"

def _style(cell, idx: int):
    bg = C_LITE if idx % 2 == 0 else C_MID
    cell.fill      = PatternFill("solid", start_color=bg, end_color=bg)
    cell.font      = Font(name="Arial", color=C_WH, size=9)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    s = Side(style="thin", color="30363D")
    cell.border    = Border(left=s, right=s, top=s, bottom=s)

def next_empty_row(ws) -> int:
    for r in range(4, 5000):
        if ws.cell(row=r, column=2).value is None:
            return r
    return ws.max_row + 1

def find_open_trade(ws, direction: str):
    """Find last row with matching direction and no exit price"""
    for r in range(ws.max_row, 3, -1):
        if ws.cell(row=r, column=5).value == direction \
           and ws.cell(row=r, column=9).value is None:
            return r
    return None

def write_entry(ws, data: dict, row: int):
    dt = parse_time(data.get("time", ""))
    vals = {
        2:  (dt.astimezone(TIMEZONE).date(), "MM/DD/YYYY"),
        4:  (get_session(dt),                None),
        5:  (data.get("direction"),          None),
        6:  (float(data.get("price",  0)),   "0.00"),
        7:  (float(data.get("sl",     0)) or None, "0.00"),
        8:  (float(data.get("tp",     0)) or None, "0.00"),
        10: (int(data.get("contracts", 1)),  "0"),
        16: (data.get("setup", "Auto"),      None),
    }
    for col, (val, fmt) in vals.items():
        c = ws.cell(row=row, column=col, value=val)
        _style(c, row - 3)
        if fmt: c.number_format = fmt
    log.info(f"Entry saved  – row {row} | {data.get('direction')} @ {data.get('price')}")

def write_exit(ws, price: float, row: int):
    c = ws.cell(row=row, column=9, value=price)
    _style(c, row - 3)
    c.number_format = "0.00"
    log.info(f"Exit saved   – row {row} | exit @ {price}")

# ═══════════════════════════════════════════════════════════════
#  ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route("/webhook", methods=["POST"])
def webhook():
    # Optional token check
    if SECRET_TOKEN and request.headers.get("X-Token", "") != SECRET_TOKEN:
        log.warning("Invalid token")
        return jsonify({"error": "Unauthorized"}), 401

    raw = request.get_data(as_text=True)
    log.info(f"Incoming webhook: {raw[:200]}")

    try:
        data = request.get_json(force=True)
    except Exception as e:
        log.error(f"JSON parse error: {e} | raw: {raw}")
        return jsonify({"error": "Invalid JSON"}), 400

    if not data:
        return jsonify({"error": "Empty payload"}), 400

    action    = str(data.get("action",    "")).lower()
    direction = str(data.get("direction", ""))
    price     = float(data.get("price",   0))

    if not os.path.exists(EXCEL_PATH):
        log.error(f"Excel not found: {EXCEL_PATH}")
        return jsonify({"error": f"Excel not found: {EXCEL_PATH}"}), 500

    try:
        wb = load_workbook(EXCEL_PATH)
        ws = wb["📋 Trade Log"]

        if action == "entry":
            row = next_empty_row(ws)
            write_entry(ws, data, row)
            wb.save(EXCEL_PATH)
            return jsonify({"status": "entry_saved", "row": row,
                            "direction": direction, "price": price}), 200

        elif action == "exit":
            row = find_open_trade(ws, direction)
            if row:
                write_exit(ws, price, row)
                wb.save(EXCEL_PATH)
                return jsonify({"status": "exit_saved", "row": row,
                                "direction": direction, "price": price}), 200
            else:
                log.warning(f"No open {direction} trade found")
                return jsonify({"status": "no_open_trade", "direction": direction}), 200

        else:
            log.warning(f"Unknown action: {action}")
            return jsonify({"status": "unknown_action", "action": action}), 200

    except Exception as e:
        log.exception("Write error")
        return jsonify({"error": str(e)}), 500


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status":     "running",
        "excel":      os.path.exists(EXCEL_PATH),
        "excel_path": EXCEL_PATH,
    }), 200


@app.route("/trades", methods=["GET"])
def list_trades():
    """Return last 20 trades as JSON – for debugging"""
    if not os.path.exists(EXCEL_PATH):
        return jsonify({"error": "Excel not found"}), 500
    wb = load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["📋 Trade Log"]
    trades = []
    for r in range(4, min(ws.max_row + 1, 24)):
        if ws.cell(row=r, column=2).value:
            trades.append({
                "row":       r,
                "date":      str(ws.cell(row=r, column=2).value),
                "direction": ws.cell(row=r, column=5).value,
                "entry":     ws.cell(row=r, column=6).value,
                "sl":        ws.cell(row=r, column=7).value,
                "tp":        ws.cell(row=r, column=8).value,
                "exit":      ws.cell(row=r, column=9).value,
                "contracts": ws.cell(row=r, column=10).value,
                "setup":     ws.cell(row=r, column=16).value,
            })
    return jsonify({"trades": trades, "count": len(trades)}), 200


# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    log.info(f"🚀 NQ Auto Journal Server starting on port {PORT}")
    log.info(f"📊 Excel file: {EXCEL_PATH}")
    app.run(host="0.0.0.0", port=PORT)
