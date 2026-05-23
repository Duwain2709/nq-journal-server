from flask import Flask, request, jsonify
import os, logging, json
from datetime import datetime

PORT = int(os.getenv("PORT", 5000))

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)

trades = []

@app.route("/webhook", methods=["POST"])
def webhook():
    raw = request.get_data(as_text=True)
    log.info(f"Incoming: {raw[:200]}")
    try:
        data = request.get_json(force=True) or {}
    except Exception as e:
        return jsonify({"error": "Invalid JSON"}), 400

    data["received_at"] = datetime.utcnow().isoformat()
    trades.append(data)
    log.info(f"Trade saved: {data}")
    return jsonify({"status": "ok", "trade": data}), 200

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "running", "trades_count": len(trades)}), 200

@app.route("/trades", methods=["GET"])
def list_trades():
    return jsonify({"trades": trades, "count": len(trades)}), 200

if __name__ == "__main__":
    log.info(f"Server starting on port {PORT}")
    app.run(host="0.0.0.0", port=PORT)
