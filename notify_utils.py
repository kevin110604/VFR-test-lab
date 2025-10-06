import requests
import os
import pytz
from datetime import datetime, timedelta
from config import UPLOAD_FOLDER

def send_teams_message(webhook_url, message):
    payload = {"text": message}
    try:
        response = requests.post(webhook_url, json=payload)
        print(f"[Teams Notify] status={response.status_code} text={response.text}")
        return response.status_code == 200
    except Exception as e:
        print("Teams webhook error:", e)
        return False

def atomic_write(filepath, text):
    tmp_path = filepath + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(text)
    os.replace(tmp_path, filepath)  # atomic trên hầu hết OS