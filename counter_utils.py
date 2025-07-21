import os
import json
import pytz
import openpyxl
from datetime import datetime, time

COUNT_JSON = "counter_stats.json"
COUNT_XLSX = "counter_log.xlsx"
DETAIL_LOG = "counter_detail_log.json"       # Nếu vẫn cần log JSON chi tiết
DETAIL_XLSX = "counter_detail_log.xlsx"
SHIFT_END = time(16, 45)  # 16:45, hết ca

# Đọc counter hiện tại (trả về dict có office, ot, date)
def read_counter():
    today_str = datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")).strftime("%Y-%m-%d")
    if os.path.exists(COUNT_JSON):
        with open(COUNT_JSON, "r", encoding="utf-8") as f:
            counter = json.load(f)
        # Nếu qua ngày mới thì reset về 0, ghi ngày mới
        if counter.get("date") != today_str:
            counter = {"date": today_str, "office": 0, "ot": 0}
    else:
        counter = {"date": today_str, "office": 0, "ot": 0}
    return counter

def write_counter_json(counter):
    with open(COUNT_JSON, "w", encoding="utf-8") as f:
        json.dump(counter, f, ensure_ascii=False)

# Ghi log tổng hợp Excel khi hết ca hoặc qua ngày
def write_counter_log_excel(counter=None, log_time=None):
    log_time = log_time or datetime.now(pytz.timezone("Asia/Ho_Chi_Minh"))
    if counter is None:
        counter = read_counter()
    if not os.path.exists(COUNT_XLSX):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Ngày", "Giờ", "Office", "OT"])
    else:
        wb = openpyxl.load_workbook(COUNT_XLSX)
        ws = wb.active
    ws.append([log_time.strftime("%Y-%m-%d"), log_time.strftime("%H:%M"), counter.get("office", 0), counter.get("ot", 0)])
     # Xoá ngày cũ nếu > 10 ngày (sau khi append)
    dates = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row+1)]
    unique_dates = []
    for d in dates:
        if d and d not in unique_dates:
            unique_dates.append(d)
    if len(unique_dates) > 10:
        # Tìm ngày cũ nhất
        remove_day = unique_dates[0]
        # Xoá tất cả dòng của ngày đó (từ dưới lên)
        for r in range(ws.max_row, 1, -1):
            if ws.cell(row=r, column=1).value == remove_day:
                ws.delete_rows(r)
    wb.save(COUNT_XLSX)

# Hàm tăng số lượng (office hoặc ot) và kiểm tra nếu đến lúc ghi log/reset thì thực hiện luôn
def update_counter(ca=None):
    now = datetime.now(pytz.timezone("Asia/Ho_Chi_Minh"))
    today_str = now.strftime("%Y-%m-%d")
    tval = now.hour * 60 + now.minute

    office_start = 8 * 60         # 8:00
    office_end = 16 * 60 + 45     # 16:45
    ot_end = 23 * 60 + 59         # 23:59

    counter = read_counter()
    # Reset nếu qua ngày mới (đảm bảo luôn đúng ngày)
    if counter.get("date") != today_str:
        # Ghi log cho ngày hôm qua nếu cần (lấy số cuối ngày hôm qua)
        write_counter_log_excel(counter)
        counter = {"date": today_str, "office": 0, "ot": 0}

    # Tự xác định ca nếu không truyền vào
    if not ca:
        if office_start <= tval < office_end:
            ca = "office"
        elif office_end <= tval <= ot_end:
            ca = "ot"
        else:
            ca = None

    if ca in ("office", "ot"):
        counter[ca] += 1

    write_counter_json(counter)

    # Nếu vừa kết thúc ca office (tại đúng 16:45)
    if tval == office_end:
        write_counter_log_excel(counter)
        # Reset counter cho ca OT
        counter = {"date": today_str, "office": 0, "ot": 0}
        write_counter_json(counter)

    return counter

# --- MỚI: GHI LOG CHI TIẾT MỖI LẦN HOÀN THÀNH REPORT RA EXCEL ---
def log_report_complete(report_number, type_of, ca):
    """Ghi log mỗi lần hoàn thành 1 report vào file Excel chi tiết."""
    now = datetime.now(pytz.timezone("Asia/Ho_Chi_Minh"))
    day_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")
    file = DETAIL_XLSX

    # Tạo file nếu chưa có
    if not os.path.exists(file):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Detail"
        ws.append(["Ngày", "Giờ", "Ca", "Report#", "Type_of"])
    else:
        wb = openpyxl.load_workbook(file)
        ws = wb.active

    ws.append([day_str, time_str, ca, report_number, type_of])
    wb.save(file)

# --- Nếu vẫn muốn log vào JSON chi tiết, giữ hàm này ---
def log_order_complete(type_of, ca):
    today_str = datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")).strftime("%Y-%m-%d")
    log_list = []
    if os.path.exists(DETAIL_LOG):
        with open(DETAIL_LOG, "r", encoding="utf-8") as f:
            log_list = json.load(f)
    log_list.append({
        "date": today_str,
        "ca": ca,  # "office" hoặc "ot"
        "type_of": type_of
    })
    with open(DETAIL_LOG, "w", encoding="utf-8") as f:
        json.dump(log_list, f, ensure_ascii=False)

# Hàm này có thể gọi khi khởi động server hoặc trước khi hiển thị home để chắc chắn log tổng hợp
def check_and_reset_counter():
    now = datetime.now(pytz.timezone("Asia/Ho_Chi_Minh"))
    today_str = now.strftime("%Y-%m-%d")
    tval = now.hour * 60 + now.minute
    office_end = 16 * 60 + 45

    counter = read_counter()
    need_log = False

    # Nếu ngày cũ → ghi log cho ngày đó rồi reset
    if counter.get("date") != today_str:
        write_counter_log_excel(counter)
        counter = {"date": today_str, "office": 0, "ot": 0}
        need_log = True

    # Nếu đã qua 16:45 và chưa log cho ca hôm nay
    if tval >= office_end and (counter.get("office", 0) > 0 or counter.get("ot", 0) > 0):
        write_counter_log_excel(counter)
        # Reset counter sau khi log
        counter = {"date": today_str, "office": 0, "ot": 0}
        need_log = True

    if need_log:
        write_counter_json(counter)
