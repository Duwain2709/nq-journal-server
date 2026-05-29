from flask import Flask, request, jsonify
from openpyxl import load_workbook
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_PATH = os.environ.get("EXCEL_PATH", "NQ_Trading_Journal.xlsx")

pending_trades = {}

def get_session(hour):
    if 1 <= hour < 8:   return "Asia"
    if 8 <= hour < 13:  return "London"
    if 13 <= hour < 15: return "NY Open"
    if 15 <= hour < 21: return "New York"
    return "After Hours"

def get_weekday(dt):
    return ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"][dt.weekday()]

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "time": datetime.utcnow().isoformat()})

@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        data = request.get_json(force=True)
        action    = data.get("action", "").lower()
        order_id  = data.get("orderId", data.get("order_id
