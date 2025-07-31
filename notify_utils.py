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

def notify_when_enough_time(
    report,
    so_gio_test,
    tag_after,
    time_file_name,
    flag_file_name,
    webhook_url,
    notify_msg,
    force_send=False,
    pending_notify_name=None
):
    from image_utils import allowed_file

    folder = os.path.join(UPLOAD_FOLDER, str(report))
    time_file = os.path.join(folder, time_file_name)
    elapsed_time = None
    start_time = None

    # Đọc timestamp bắt đầu test (ảnh before)
    if os.path.exists(time_file):
        with open(time_file, "r", encoding="utf-8") as f:
            tstr = f.read().strip()
        try:
            start_time = datetime.strptime(tstr, "%d/%m/%Y %H:%M")
            now = datetime.now()
            elapsed_time = (now - start_time).total_seconds() / 3600
        except Exception:
            start_time = None
            elapsed_time = None

    # Kiểm tra đã có ảnh after chưa
    after_img_exists = False
    if os.path.exists(folder):
        for f in os.listdir(folder):
            if allowed_file(f) and f.startswith(tag_after):
                after_img_exists = True
                break

    enough_time = (elapsed_time is not None and elapsed_time >= so_gio_test)
    sent = False

    flag_path = os.path.join(folder, flag_file_name) if flag_file_name else None

    # Kiểm tra thời gian hiện tại (giờ local)
    now = datetime.now()
    cur_hour = now.hour
    ALLOWED_HOUR_START = 8
    ALLOWED_HOUR_END = 21

    # Xử lý gửi ngay khi đủ giờ, hoặc pending nếu ngoài khung giờ
    if enough_time and not after_img_exists:
        send_now = ALLOWED_HOUR_START <= cur_hour < ALLOWED_HOUR_END
        # Nếu ngoài giờ gửi thì tạo pending
        if not send_now:
            if pending_notify_name:
                pending_path = os.path.join(folder, pending_notify_name)
                atomic_write(pending_path, notify_msg)
            return {"show_notice": True, "sent": False}
        # Gửi bình thường nếu trong giờ cho phép
        if force_send or (flag_file_name is None) or (flag_file_name and not os.path.exists(flag_path)):
            send_teams_message(webhook_url, notify_msg)
            sent = True
            # Đánh dấu đã gửi lần đầu nếu dùng flag
            if flag_path and not os.path.exists(flag_path):
                atomic_write(flag_path, now.strftime("%Y-%m-%d %H:%M:%S"))
            # Xóa pending notify nếu có
            if pending_notify_name:
                pending_path = os.path.join(folder, pending_notify_name)
                if os.path.exists(pending_path):
                    try:
                        os.remove(pending_path)
                    except Exception:
                        pass  # Có thể file đã bị xóa bởi request khác, không sao

    return {"show_notice": enough_time and not after_img_exists, "sent": sent}

