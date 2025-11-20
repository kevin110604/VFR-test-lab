from flask import Flask, request, render_template, session, redirect, url_for, jsonify, flash, send_from_directory, Response, stream_with_context, abort, template_rendered, send_file
from config import SECRET_KEY, local_main, UPLOAD_FOLDER, TEST_GROUPS, local_complete, SO_GIO_TEST, TEAMS_WEBHOOK_URL_TRF, TEAMS_WEBHOOK_URL_RATE, TEMPLATE_MAP
from excel_utils import get_item_code, get_col_idx, copy_row_with_style, write_tfr_to_excel, append_row_to_trf, export_expired_samples_to_excel

from image_utils import allowed_file, get_img_urls
from auth import login, get_user_type
from test_logic import load_group_notes, get_group_test_status, is_drop_test, is_impact_test, is_rotational_test,  TEST_GROUP_TITLES, TEST_TYPE_VI, DROP_ZONES, DROP_LABELS, GT68_FACE_LABELS, GT68_FACE_ZONES
from test_logic import IMPACT_ZONES, IMPACT_LABELS, ROT_LABELS, ROT_ZONES, RH_IMPACT_ZONES, RH_VIB_ZONES, RH_SECOND_IMPACT_ZONES, RH_STEP12_ZONES, TWO_C_NP_STEP5_ZONES, update_group_note_file, get_group_note_value, F2057_TEST_TITLES
from notify_utils import send_teams_message
from counter_utils import update_counter, check_and_reset_counter, log_report_complete
from docx_utils import approve_request_fill_docx_pdf, fill_cover_from_excel_generic, try_convert_to_pdf
from file_utils import (
    safe_write_json, safe_read_json, safe_save_excel, safe_load_excel,
    safe_write_text, safe_read_text, safe_append_backup_json   # <— thêm hàm này
)
import re, os, pytz, json, openpyxl, random, subprocess, regex, traceback, calendar, time, tempfile, uuid, secrets, copy, glob, zipfile, io
from contextlib import contextmanager
from datetime import datetime, timedelta
from waitress import serve
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from collections import defaultdict, OrderedDict
from threading import Lock
from contextlib import contextmanager
from vfr3 import vfr3_bp
from werkzeug.utils import secure_filename
from qr_print import qr_bp
from testlab_dashboard import dashboard_bp
from flask_session import Session
from collections import OrderedDict
from markupsafe import Markup
from dateutil.relativedelta import relativedelta
from datetime import datetime, date
from uuid import uuid4
from io import BytesIO

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.register_blueprint(vfr3_bp)
app.register_blueprint(qr_bp)
app.register_blueprint(dashboard_bp, url_prefix="/testlab")

# ==== Server-side session config ====
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config.update(
    SESSION_TYPE='filesystem',  # có thể đổi sang 'redis' / 'sqlalchemy' nếu muốn
    SESSION_FILE_DIR=os.path.join(BASE_DIR, '.flask_session'),
    SESSION_PERMANENT=False,  # cookie phiên (đóng trình duyệt là hết)
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
    SESSION_COOKIE_NAME='vfr_session',
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',  # đổi 'Strict' nếu kiosk
    # SESSION_COOKIE_SECURE=True,   # bật nếu chạy HTTPS
)
Session(app)
# Những test dùng giao diện Hot & Cold
HOTCOLD_LIKE = set(["hot_cold", "standing_water", "stain", "corrosion"])
INDOOR_GROUPS = {"indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal","outdoor_finishing"}
REPORT_NO_LOCK = Lock()
BLANK_TOKENS = {"", "-", "—"}
BOXES_FILE = os.path.join(os.path.dirname(__file__), "boxes.json")

_LAST_TEST_LRU_LIMIT = 50  # lưu tối đa 50 report gần nhất

def read_boxes():
    data = safe_read_json(BOXES_FILE)
    if isinstance(data, list):
        return data
    if isinstance(data, dict) and isinstance(data.get("boxes"), list):
        return data["boxes"]
    return ["B1-S1","B1-S2","B2-S1","B2-S2","B3-S1"]

def write_boxes(boxes):
    safe_write_json(BOXES_FILE, {"boxes": boxes})

def _set_limited_mapping(session_key: str, subkey: str, value, limit: int = _LAST_TEST_LRU_LIMIT):
    """
    Lưu value vào session[session_key][subkey] với giới hạn số phần tử.
    Dùng OrderedDict như LRU: nếu quá 'limit' thì pop phần tử cũ nhất.
    """
    mapping = session.get(session_key)
    if not isinstance(mapping, dict):
        mapping = {}

    od = OrderedDict(mapping)
    if subkey in od:
        od.move_to_end(subkey)
    od[subkey] = value

    while len(od) > limit:
        od.popitem(last=False)

    session[session_key] = dict(od)

def set_last_test(report: str, group_code: str, display_title: str = None):
    """
    Ghi đồng thời 'last_test_code' (group) và 'last_test_type' (display title).
    Nếu không truyền display_title sẽ dùng tên nhóm (VI) qua get_group_title.
    """
    if display_title is None:
        display_title = get_group_title(group_code) or group_code.upper()
    _set_limited_mapping("last_test_code", str(report), group_code)
    _set_limited_mapping("last_test_type", str(report), display_title)

def set_last_test_type(report: str, display_title: str):
    """Chỉ cập nhật 'last_test_type' (không đụng tới last_test_code)."""
    _set_limited_mapping("last_test_type", str(report), display_title)

def get_last_test_code(report: str, default=None):
    return (session.get("last_test_code") or {}).get(str(report), default)

def get_last_test_type(report: str, default=None):
    return (session.get("last_test_type") or {}).get(str(report), default)

def _is_blank_cell(v):
    if v is None:
        return True
    if isinstance(v, str):
        s = (v.replace("\u00A0","").replace("\u200B","")
               .replace("\r","").replace("\n","").replace("\t","").strip())
        return s in BLANK_TOKENS or s == ""
    return False

def row_is_filled_for_report(excel_path, report_no):
    """True nếu dòng có B == report_no ĐÃ có dữ liệu ở bất kỳ cột C..X; False nếu vẫn trống."""
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active
    target_row = None
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value  # cột B
        if (str(v).strip() if v is not None else "") == str(report_no).strip():
            target_row = r
            break
    if target_row is None:
        wb.close()
        # Không thấy mã trong cột B (khác thiết kế) -> coi như đã dùng để tránh ghi bậy
        return True
    for c in range(3, 25):  # C..X
        if not _is_blank_cell(ws.cell(row=target_row, column=c).value):
            wb.close()
            return True   # ĐÃ có dữ liệu
    wb.close()
    return False          # C..X đều trống => CHƯA dùng

def format_excel_date_short(dt):
    """Convert Python datetime/date -> format 'd-mmm' (e.g., 7-Aug) cho Excel."""
    if isinstance(dt, str):
        # Thử parse về date
        try:
            dt = datetime.strptime(dt, "%Y-%m-%d")
        except:
            try:
                dt = datetime.strptime(dt, "%d/%m/%Y")
            except:
                try:
                    dt = datetime.strptime(dt, "%m/%d/%Y")
                except:
                    return dt  # Trả nguyên nếu không parse được
    # Trả về dạng 'd-mmm'
    return f"{dt.day}-{calendar.month_abbr[dt.month]}"

def try_parse_excel_date(dt):
    """Parse dt về kiểu datetime/date nếu có thể, trả về None nếu không hợp lệ."""
    if isinstance(dt, datetime):
        return dt
    if isinstance(dt, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(dt, fmt)
            except Exception:
                continue
    return None

def calculate_default_etd(request_date: str, test_group: str, *, all_reqs=None) -> str:
    if not request_date:
        return ""

    # Chuẩn hoá group
    g = _group_of(test_group)
    if g in ("CONSTRUCTION", "TRANSIT"):
        base = 2   # 3 ngày tính cả ngày request => +2
    elif g in ("FINISHING", "MATERIAL"):
        base = 4   # 5 ngày tính cả ngày request => +4
    else:
        base = 2

    # --- Chuẩn bị dữ liệu Pending + Archive ---
    # Pending: chỉ lấy Submitted
    pending_submitted = []
    if isinstance(all_reqs, list):
        for r in all_reqs:
            try:
                if (r.get("status") or "").strip() == "Submitted":
                    pending_submitted.append(r)
            except Exception:
                continue

    # Archive: dùng trực tiếp test_group nếu có, fallback map từ report_no
    try:
        archive_list = safe_read_json(ARCHIVE_LOG) or []
    except Exception:
        archive_list = []

    rep2grp = _build_reportno_to_group_map()

    archive_mapped = []
    for a in archive_list:
        try:
            req_date = (a.get("request_date") or "").strip()
            trq     = (a.get("trq_id") or "").strip()
            rep_no  = (a.get("report_no") or "").strip()
            grp0    = (a.get("test_group") or "").strip()  # ưu tiên group đã lưu
            grp     = grp0 if grp0 else rep2grp.get(rep_no, "")
            if not req_date or not grp:
                continue
            archive_mapped.append({
                "request_date": req_date,
                "test_group": grp,       # để _group_of dùng
                "trq_id": trq or rep_no, # fallback report_no nếu thiếu trq_id
                "status": "Approved",
            })
        except Exception:
            continue

    # --- Đếm TRQ duy nhất cho (request_date, group) ---
    target_date = (request_date or "").strip()
    target_grp  = g
    uniq_trq = set()

    # Pending Submitted
    for r in pending_submitted:
        try:
            r_date = (r.get("request_date") or "").strip()
            r_grp  = _group_of(r.get("test_group") or r.get("type_of_test"))
            if r_date == target_date and r_grp == target_grp:
                tid = (r.get("trq_id") or "").strip()
                if not tid:
                    # để không crash, coi 1 dòng là 1 "TRQ"
                    tid = f"__row_{id(r)}"
                uniq_trq.add(tid)
        except Exception:
            continue

    # Archive Approved (đã map group & trq)
    for a in archive_mapped:
        try:
            if a["request_date"] == target_date and _group_of(a["test_group"]) == target_grp:
                tid = (a.get("trq_id") or "").strip()
                if not tid:
                    tid = a.get("report_no", "")
                if tid:
                    uniq_trq.add(tid)
        except Exception:
            continue

    cnt = len(uniq_trq)  # số TRQ duy nhất đã có TRƯỚC request mới

    # --- Extra theo ngưỡng từng nhóm ---
    extra = 0
    if g in ("CONSTRUCTION", "TRANSIT"):
        # đang là # (cnt + 1); nếu đã có ≥5 thì request mới rơi vào #6..#10
        if cnt >= 10:      # đang là #11..#15
            extra = 2
        elif cnt >= 5:     # đang là #6..#10
            extra = 1
    elif g in ("FINISHING", "MATERIAL"):
        if cnt >= 30:      # đang là #31..#45
            extra = 4
        elif cnt >= 15:    # đang là #16..#30
            extra = 2

    try:
        d0 = datetime.strptime(request_date, "%Y-%m-%d").date()
    except Exception:
        try:
            d0 = datetime.strptime(request_date, "%d/%m/%Y").date()
        except Exception:
            return ""

    etd = d0 + timedelta(days=base + extra)
    return etd.strftime("%Y-%m-%d")

# ---- các hàm helper không đổi (giữ nguyên) ----
def get_group_title(group):
    for g_id, g_name in TEST_GROUPS:
        if g_id == group:
            return g_name
    return None

def generate_unique_trq_id(existing_ids):
    yy = str(datetime.now().year)[-2:]  # 2 số cuối của năm hiện tại
    while True:
        num = random.randint(10000, 99999)
        new_id = f"TL-{yy}{num}"
        if new_id not in existing_ids:
            return new_id

ARCHIVE_LOG = "tfr_archive.json"
TFR_LOG_FILE = "tfr_requests.json"

@contextmanager
def report_lock():
    lock_dir = tempfile.gettempdir()              # Windows: C:\Users\<user>\AppData\Local\Temp
    lock_path = os.path.join(lock_dir, "tfr_report.lock")
    # Optional: timeout để không chờ vô hạn
    timeout_s = 60
    t0 = time.time()
    fd = None
    while True:
        try:
            # tạo mới, nếu đã có -> FileExistsError
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
            # ghi chút info để debug stale lock
            os.write(fd, f"pid={os.getpid()} run={uuid.uuid4()}".encode("utf-8"))
            break
        except FileExistsError:
            # lock lâu quá coi như stale => cố gắng xoá
            if time.time() - t0 > timeout_s:
                _try_unlink_with_retry(lock_path)
            else:
                time.sleep(0.05 + random.random() * 0.15)
    try:
        yield
    finally:
        try:
            if fd is not None:
                os.close(fd)
        except Exception:
            pass
        _try_unlink_with_retry(lock_path)

def _try_unlink_with_retry(path, retries=8, delay=0.08):
    # Windows có thể vướng PermissionError do AV; retry ngắn sẽ qua được
    for i in range(retries):
        try:
            os.unlink(path)
            return True
        except FileNotFoundError:
            return True
        except PermissionError:
            time.sleep(delay * (1.5 ** i))  # backoff
        except Exception:
            time.sleep(delay)
    # fallback: đổi tên để không cản trở lần sau
    try:
        os.rename(path, path + ".stale")
    except Exception:
        pass
    return False

def bump_report_no(s):
    m = re.search(r'(\d+)$', str(s))
    if not m:
        return f"{s}-1"
    start, end = m.span(1)
    n = int(m.group(1)) + 1
    width = end - start
    return f"{s[:start]}{str(n).zfill(width)}"

def report_no_exists(report_no, tfr_requests):
    """
    ĐÃ DÙNG khi:
    - Dòng B==report_no trong Excel có dữ liệu C..X (không còn trống), HOẶC
    - File đầu ra cho mã đó đã tồn tại (pdf/docx), HOẶC
    - Mã này đã nằm trong archive/log (đã approve).
    """
    # 1) Excel: dòng đã có dữ liệu?
    try:
        if row_is_filled_for_report(local_main, report_no):
            return True
    except Exception:
        pass

    # 2) Trùng file đã sinh?
    output_folder = os.path.join('static', 'TFR')
    if os.path.exists(os.path.join(output_folder, f"{report_no}.pdf")):
        return True
    if os.path.exists(os.path.join(output_folder, f"{report_no}.docx")):
        return True

    # 3) Trùng trong log pending đang dùng?
    for r in tfr_requests:
        if str(r.get("report_no") or "").strip() == str(report_no):
            return True

    # 4) Trùng trong archive (đã approve)?
    try:
        archive = safe_read_json(ARCHIVE_LOG)
        for r in archive:
            if str(r.get("report_no") or "").strip() == str(report_no):
                return True
    except Exception:
        pass

    return False

def allocate_unique_report_no(make_report_func, req, tfr_requests, max_try=2):
    """
    Cấp và cố định report_no đúng logic:
    - Nếu req đã có report_no: kiểm tra dòng B==report_no còn trống (C..X). Nếu đã có dữ liệu -> báo lỗi.
    - Nếu chưa có: để make_report_func chọn DÒNG TRỐNG (C..X trống) và trả về report_no tương ứng.
    - Có retry nhẹ (bump số) nếu hi hữu bị chiếm chỗ giữa chừng.
    """
    tries = 0

    # Case A: preset report_no trong req
    fixed_req = dict(req)
    preset = str(fixed_req.get("report_no", "")).strip()
    if preset:
        if row_is_filled_for_report(local_main, preset):
            raise RuntimeError(f"Mã report {preset} đã có dữ liệu, không thể ghi đè.")
        pdf_path, report_no = make_report_func(fixed_req)  # docx_utils ưu tiên số đã set
        return pdf_path, report_no

    # Case B: chưa có -> để make_report_func chọn dòng C..X trống
    while True:
        pdf_path, report_no = make_report_func(req)

        # Xác nhận lại: dòng vẫn còn trống?
        if not row_is_filled_for_report(local_main, report_no):
            return pdf_path, report_no

        # Hi hữu: ai đó vừa điền vào dòng này giữa chừng -> thử lại (bump số + regenerate)
        tries += 1
        if tries >= max_try:
            raise RuntimeError("Không tìm được dòng trống để cấp mã report.")

        # Xoá file vừa sinh (đi nhầm dòng)
        try:
            outdir = os.path.join('static', 'TFR')
            for ext in ('.pdf', '.docx'):
                fp = os.path.join(outdir, f"{report_no}{ext}")
                if os.path.exists(fp):
                    os.remove(fp)
        except Exception:
            pass

        # Bump số và tái tạo với số cố định
        tries += 1
        if tries >= max_try:
            raise RuntimeError("Không cấp được report_no duy nhất sau nhiều lần thử")

        bumped = bump_report_no(report_no)
        fixed_req = dict(req)
        fixed_req["report_no"] = bumped
        pdf_path, report_no = make_report_func(fixed_req)

# ---- ARCHIVE REQUEST LOG ----
def archive_request(short_data):
    now = datetime.now()
    archive = safe_read_json(ARCHIVE_LOG)
    def get_dt(d):
        if "-" in d: return datetime.strptime(d, "%Y-%m-%d")
        else: return datetime.strptime(d, "%d/%m/%Y")
    archive = [r for r in archive if (now - get_dt(r["request_date"])).days < 14]
    archive.append(short_data)
    safe_write_json(ARCHIVE_LOG, archive)
    safe_append_backup_json(ARCHIVE_LOG, short_data) 

# --- ADD NEW: cleanup archive file (>14 ngày) ---
def cleanup_archive_json(days=14):
    """
    Xóa các bản ghi archive quá 'days' ngày (xóa thật trong JSON).
    Ưu tiên ARCHIVE_LOG / TFR_ARCHIVE_FILE nếu có; nếu không suy ra từ TFR_LOG_FILE.
    """
    try:
        archive_path = globals().get("ARCHIVE_LOG") or globals().get("TFR_ARCHIVE_FILE")
        if not archive_path:
            base, ext = os.path.splitext(TFR_LOG_FILE)
            archive_path = f"{base}_archive.json"

        data = safe_read_json(archive_path)
        if not isinstance(data, list) or not data:
            return
        def _parse_date(s):
            if not s:
                return None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
                try:
                    return datetime.strptime(str(s), fmt).date()
                except Exception:
                    pass
            return None

        today = datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")).date()
        kept = []
        for r in data:
            d = None
            if isinstance(r, dict):
                d = _parse_date(r.get("approved_date")) or _parse_date(r.get("etd")) or _parse_date(r.get("request_date"))
            if not d or (today - d).days <= days:
                kept.append(r)
        if len(kept) != len(data):
            safe_write_json(archive_path, kept)
    except Exception as _e:
        print("cleanup_archive_json error:", _e)

# ---- HOME PAGE ----
@app.route("/", methods=["GET", "POST"])
def home():
    message = None

    # Handle item# search and login/set_staff_id POST actions
    if request.method == "POST":
        item_search = request.form.get("item_search", "").strip()
        if item_search:
            return redirect(url_for("home", item_search=item_search))
        if request.form.get("action") == "login":
            password_input = request.form.get("password", "")
            if login(password_input):
                # auth.login đã set session['auth_ok'] và session['user_type']
                session['role'] = get_user_type()  # 'stl' / 'wtl' / 'vfr3'
                return redirect(url_for("home"))
            else:
                message = "Incorrect password. Please try again."
        elif request.form.get("action") == "set_staff_id":
            staff_id = request.form.get("staff_id", "").strip()
            # Regex: số - họ tên (hỗ trợ Unicode tiếng Việt)
            pattern = r'^\d+\s*-\s*[\p{L}]+(?:\s+[\p{L}]+){1,}$'
            if not staff_id:
                message = "Please enter your ID!"
            elif not regex.match(pattern, staff_id):
                message = ("Please enter the correct format: Staff ID - Full name "
                          "(e.g., 19797 - Nguyen Van A)")
            else:
                session["staff_id"] = staff_id
                return redirect(url_for("home"))


    # ==== Load Excel data ====
    full_report_list = []
    type_of_set = set()
    all_statuses = ['LATE', 'MUST', 'DUE', 'ACTIVE', 'COMPLETE', 'DONE']
    raw_status_set = set()
    try:
        wb = load_workbook(local_main)
        ws = wb.active
        def clean_col(s):
            s = str(s).lower().strip()
            s = re.sub(r'[^a-z0-9#]+', '', s)
            return s
        headers = {}
        for col in range(2, ws.max_column + 1):
            name = ws.cell(row=1, column=col).value
            if name:
                clean = clean_col(name)
                headers[clean] = col
        report_col    = headers.get("report#")
        item_col      = headers.get("item#")
        status_col    = headers.get("status")
        test_date_col = headers.get("logindate")
        type_of_col   = headers.get("typeof")
        etd_col       = headers.get("etd")
        if None in (report_col, item_col, status_col, test_date_col):
            message = f"Missing columns in Excel file! Found: {headers}"
        else:
            for row in range(2, ws.max_row + 1):
                status_raw = ws.cell(row=row, column=status_col).value
                status = str(status_raw).strip().upper() if status_raw else ""
                report = ws.cell(row=row, column=report_col).value
                item = ws.cell(row=row, column=item_col).value
                etd = ws.cell(row=row, column=etd_col).value if etd_col else ""
                type_of = ws.cell(row=row, column=type_of_col).value if type_of_col else ""
                log_date = ws.cell(row=row, column=test_date_col).value
                log_date_str = str(log_date).strip() if log_date else ""
                if log_date_str: log_date_str = log_date_str.split()[0]
                r_dict = {
                    "report": str(report).strip() if report else "",
                    "item": str(item).strip() if item else "",
                    "status": status,
                    "type_of": str(type_of).strip() if type_of else "",
                    "log_date": log_date_str,
                    "etd": etd if etd is not None else ""
                }
                full_report_list.append(r_dict)
                if r_dict["type_of"]:
                    type_of_set.add(r_dict["type_of"])
                if status:
                    raw_status_set.add(status)
        type_of_set = sorted(type_of_set)
    except Exception as e:
        message = f"Error reading list: {e}"

    status_set = []
    all_statuses = ['LATE', 'MUST', 'DUE', 'ACTIVE', 'COMPLETE', 'DONE']
    for s in all_statuses:
        if s in raw_status_set: status_set.append(s)
    for s in sorted(raw_status_set):
        if s not in status_set: status_set.append(s)

    # --- LOGIC LỌC STATUS ---
    selected_status = request.args.getlist("status")
    filter_from_user = "status" in request.args

    if not filter_from_user:
        # Mới vào trang, mặc định lọc theo LATE, MUST, DUE
        selected_status = ["LATE", "MUST", "DUE"]
    else:
        # Nếu form lọc được gửi (dù bấm All hay chọn từng status)
        # Nếu không chọn gì hoặc chỉ tick All → ALL (không filter theo status)
        if not selected_status or selected_status == [""]:
            selected_status = []
        elif "" in selected_status:
            # Nếu có tick cả All + các status khác, vẫn xem như ALL
            selected_status = []

    selected_type = request.args.get("type_of", "")
    item_search = request.args.get("item_search", "").strip()

    report_list = full_report_list
    if item_search:
        # Khi tìm item thì luôn tìm trên toàn bộ danh sách, không lọc theo trạng thái!
        report_list = [r for r in full_report_list if item_search.lower() in (r["item"] or "").lower()]
        if selected_type:
            report_list = [r for r in report_list if r["type_of"] == selected_type]
        def safe_report_key(r):
            try:
                return int(r["report"])
            except:
                return r["report"]
        report_list = sorted(report_list, key=safe_report_key)
    else:
        if selected_type:
            report_list = [r for r in report_list if r["type_of"] == selected_type]
        if selected_status:
            report_list = [r for r in report_list if r["status"] in selected_status]

    try:
        page = int(request.args.get("page", "1"))
    except:
        page = 1
    try:
        page_size = int(request.args.get("page_size", "10"))
    except:
        page_size = 10
    if page_size not in [10, 15, 20]: page_size = 10
    total_reports = len(report_list)
    total_pages = max((total_reports + page_size - 1) // page_size, 1)
    if page < 1: page = 1
    if page > total_pages: page = total_pages
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    report_list_page = report_list[start_idx:end_idx]

    type_shortname = {
        "CONSTRUCTION": "CON",
        "FINISHING": "FIN",
        "MATERIAL": "MAT",
        "PACKING": "PAC",
        "GENERAL": "GEN",
    }
    summary_by_type = []
    for t in type_of_set:
        late   = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "LATE")
        due    = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "DUE")
        must   = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "MUST")
        active = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "ACTIVE")
        total  = late + due + must + active

        t_up = (t or "").upper()
        if t_up.startswith("OUTSOURCE"):
            # Lấy 3 ký tự cuối sau dấu '-' nếu có; nếu không có '-' thì lấy 3 ký tự cuối của toàn chuỗi
            parts = t_up.split("-")
            short = (parts[-1][-3:] if len(parts) >= 2 else t_up[-3:])
        else:
            short = type_shortname.get(t, t_up[:3])

        summary_by_type.append({
            "type_of": t,
            "short": short,
            "late": late,
            "due": due,
            "must": must,
            "active": active,
            "total": total,
        })

    counter = {"office": 0, "ot": 0}
    path = "counter_stats.json"
    if os.path.exists(path):
        counter = safe_read_json(path)

    return render_template(
        "home.html",
        message=message,
        type_of_set=type_of_set,
        selected_type=selected_type,
        status_set=status_set,
        selected_status=selected_status,
        session=session,
        report_list=report_list,
        summary_by_type=summary_by_type,
        counter=counter,
        page=page,
        total_pages=total_pages,
        page_size=page_size,
        total_reports=total_reports,
        request=request,
        darkmode=session.get('darkmode', '0'),
        slang=session.get('lang', 'vi'),
    )

def get_category_component_position(finishing_type, material_type):
    # material_type: chỉ nhận WOOD hoặc METAL (nên xử lý hoa thường hóa)
    if not finishing_type or not material_type:
        return ""
    finishing_type = finishing_type.strip().upper()
    material_type = material_type.strip().upper()
    if finishing_type == "QA TEST":
        if material_type == "WOOD":
            return "COLOR PANEL"
        elif material_type == "METAL":
            return "METAL"
    elif finishing_type == "LINE TEST":
        if material_type == "WOOD":
            return "LINE TEST_COLOR"
        elif material_type == "METAL":
            return "LINE TEST_METAL"
    return ""

def _load_pending_locked():
    with PENDING_LOCK:
        return safe_read_json(TFR_LOG_FILE)

def _write_pending_locked(data):
    with PENDING_LOCK:
        safe_write_json(TFR_LOG_FILE, data)

# ==== Helpers nhóm test ====
def _group_of(test_group: str) -> str:
    """
    Chuẩn hoá nhóm test để tính ETD: CONSTRUCTION / TRANSIT / FINISHING / MATERIAL
    """
    g = (test_group or "").strip().upper()
    if "CONSTRUCTION" in g: return "CONSTRUCTION"
    if "TRANSIT" in g:      return "TRANSIT"
    if "FINISHING" in g:    return "FINISHING"
    if "MATERIAL" in g:     return "MATERIAL"
    return g or "OTHER"

def compute_request_date_now(cutoff_hour: int = 15) -> str:
    """
    Quy tắc request_date:
    - Trước 15:00  -> hôm nay
    - Từ 15:00 trở đi -> ngày mai
    """
    now = datetime.now()
    today = now.date()
    if now.hour >= cutoff_hour:
        return (today + timedelta(days=1)).strftime("%Y-%m-%d")
    return today.strftime("%Y-%m-%d")

def _count_by_date_and_group(all_reqs, req_date: str, group_name: str) -> int:
    """
    Đếm số request theo (request_date, group) trên 1 danh sách 'all_reqs'.
    Đếm THEO REQUEST (mỗi record = 1), và chỉ tính các record có status != Declined
    (với pending), nhằm phục vụ tính ETD.
    """
    gn = _group_of(group_name)
    dd = (req_date or "").strip()
    c = 0
    for r in (all_reqs or []):
        try:
            r_date = (r.get("request_date") or "").strip()
            r_group = _group_of(r.get("test_group") or r.get("type_of_test"))
            if r_date == dd and r_group == gn:
                st = (r.get("status") or "").strip()
                # Pending: chỉ tính Submitted; nếu record đến từ archive có thể không có status -> vẫn tính
                if st and st != "Submitted" and st != "Approved":
                    # loại Declined và các trạng thái pending khác
                    continue
                c += 1
        except Exception:
            continue
    return c


def _build_reportno_to_group_map():
    """
    Dò Excel 'local_main' để map report_no -> group chuẩn hoá (CONSTRUCTION/TRANSIT/FINISHING/MATERIAL/OTHER)
    Dựa vào cột 'type of' mà approve_all_one() đã điền.
    """
    try:
        wb = safe_load_excel(local_main)
        ws = wb.active
        col_report = get_col_idx(ws, "report#")
        col_typeof = get_col_idx(ws, "type of")
        if not col_report or not col_typeof:
            return {}

        mapping = {}
        for row in range(2, ws.max_row + 1):
            rep = ws.cell(row=row, column=col_report).value
            tp  = ws.cell(row=row, column=col_typeof).value
            rep_s = ("" if rep is None else str(rep)).strip()
            tp_s  = ("" if tp  is None else str(tp )).strip()
            if not rep_s:
                continue
            # Excel lưu "type of" KHÔNG có " TEST", nên thêm " TEST" để _group_of() hiểu,
            # hoặc bạn có thể map thẳng nếu muốn.
            grp = _group_of(tp_s + " TEST") if tp_s else "OTHER"
            mapping[rep_s] = grp
        return mapping
    except Exception:
        return {}


def calculate_default_etd(request_date: str, test_group: str, *, all_reqs=None) -> str:
    """
    ETD mặc định, tính từ request_date (tính CẢ ngày request).
    TẢI TRONG NGÀY = Pending (chỉ Submitted) + Approved (Archive), loại bỏ Declined.

    - CONSTRUCTION / TRANSIT: 3 ngày  -> base +2 ngày
      * tải (trong cùng request_date), đếm THEO REQUEST:
        - đã có ≥5  req  (đang là req #6..#10)  -> +1 ngày
        - đã có ≥10 req (đang là req #11..#15) -> +2 ngày

    - FINISHING / MATERIAL : 5 ngày  -> base +4 ngày
      * tải:
        - đã có ≥15 req (đang là #16..#30) -> +2 ngày
        - đã có ≥30 req (đang là #31..#45) -> +4 ngày
    """
    if not request_date:
        return ""

    g = _group_of(test_group)
    if g in ("CONSTRUCTION", "TRANSIT"):
        base = 2   # 3 ngày tính cả ngày request => +2
    elif g in ("FINISHING", "MATERIAL"):
        base = 4   # 5 ngày tính cả ngày request => +4
    else:
        base = 2

    # GHÉP Pending (chỉ Submitted) + Archive để đếm tải theo ngày/type
    try:
        archive_list = safe_read_json(ARCHIVE_LOG) or []
    except Exception:
        archive_list = []

    # Archive không có test_group, nên join qua Excel để gán group
    rep2grp = _build_reportno_to_group_map()
    archive_mapped = []
    for a in archive_list:
        try:
            # giữ các khóa cần cho _count_by_date_and_group
            req_date = (a.get("request_date") or "").strip()
            rep_no   = (a.get("report_no") or "").strip()
            grp      = rep2grp.get(rep_no, "")
            if not req_date or not grp:
                continue
            archive_mapped.append({
                "request_date": req_date,
                "test_group": grp,       # để _group_of hiểu
                "status": "Approved",    # đánh dấu để lọc hợp lệ
            })
        except Exception:
            continue

    combined = []
    # Pending (all_reqs) – chỉ muốn Submitted
    if isinstance(all_reqs, list):
        combined += [r for r in all_reqs if (r.get("status") or "").strip() == "Submitted"]
    # Approved (archive đã map)
    combined += archive_mapped

    cnt = _count_by_date_and_group(combined, request_date, g)

    extra = 0
    if g in ("CONSTRUCTION", "TRANSIT"):
        if cnt >= 10:      # đang là #11..#15
            extra = 2
        elif cnt >= 5:     # đang là #6..#10
            extra = 1
    elif g in ("FINISHING", "MATERIAL"):
        if cnt >= 30:      # đang là #31..#45
            extra = 4
        elif cnt >= 15:    # đang là #16..#30
            extra = 2

    d0 = datetime.strptime(request_date, "%Y-%m-%d").date()
    etd = d0 + timedelta(days=base + extra)
    return etd.strftime("%Y-%m-%d")

TFR_INIT_DIR = os.path.join('static', 'TFR_INIT')
os.makedirs(TFR_INIT_DIR, exist_ok=True)

def _save_initial_img(file_storage, trq_id):
    """Lưu ảnh initial theo TRQ-ID, trả về đường dẫn tương đối dưới /static (ví dụ: 'TFR_INIT/TRQ123_20250101_120102.jpg')."""
    if not file_storage or not getattr(file_storage, 'filename', ''):
        return None
    fname = secure_filename(file_storage.filename)
    ext = (fname.rsplit('.', 1)[-1] if '.' in fname else 'jpg').lower()
    if not allowed_file(fname):
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{trq_id}_{stamp}.{ext}"
    abs_path = os.path.join(TFR_INIT_DIR, out_name)
    file_storage.save(abs_path)
    return f"TFR_INIT/{out_name}"

@app.route("/tfr_request_form", methods=["GET", "POST"])
def tfr_request_form():
    tfr_requests = _load_pending_locked()
    error = ""
    form_data = {}
    missing_fields = []

    # Lấy tham số từ URL hoặc từ POST
    trq_id = request.args.get("trq_id") or request.form.get("trq_id")
    edit_idx = request.args.get("edit_idx") or request.form.get("edit_idx")
    editing = False

    # Nếu có trq_id + edit_idx -> đang ở chế độ EDIT: nạp sẵn dữ liệu vào form_data
    if trq_id:
        try:
            # Lấy tất cả vị trí có cùng TRQ-ID trong file gốc
            matches = [i for i, req in enumerate(tfr_requests) if (req.get("trq_id") or "").strip() == str(trq_id).strip()]

            if matches:
                # Nếu có nhiều bản ghi cùng TRQ (trường hợp admin giữ TRQ khi duplicate)
                # và edit_idx POST lên là ordinal trong 'matches' thì dùng, ngược lại lấy phần tử đầu tiên.
                sel = 0
                if edit_idx is not None:
                    try:
                        _ordinal = int(edit_idx)
                        if 0 <= _ordinal < len(matches):
                            sel = _ordinal
                    except Exception:
                        pass

                abs_idx = matches[sel]
                form_data = tfr_requests[abs_idx].copy()
                editing = True
                # Đảm bảo hidden edit_idx là "chỉ số tuyệt đối" để các lần submit sau không lệch
                form_data["edit_idx"] = str(abs_idx)
            else:
                # Không tìm thấy theo TRQ-ID -> coi như tạo mới
                editing = False
        except Exception:
            editing = False

    if request.method == "POST":
        form = request.form

        required_fields = [
            "requestor", "employee_id", "department", "request_date",
            "sample_description", "test_status", "quantity", "sample_return"
        ]
        for field in required_fields:
            if not form.get(field) and not form.get(f"{field}_na"):
                missing_fields.append(field)

        test_group = form.get("test_group", "")
        if not test_group:
            missing_fields.append("test_group")
            error = "Phải chọn loại test!"

        furniture_testing = form.get("furniture_testing", "")
        if not furniture_testing:
            missing_fields.append("furniture_testing")
            error = "Phải chọn Indoor hoặc Outdoor!"

        finishing_type = form.get("finishing_type", "")
        material_type = form.get("material_type", "")

        # form_data để render lại khi lỗi
        form_data = form.to_dict(flat=True)
        form_data["test_group"] = test_group
        form_data["furniture_testing"] = furniture_testing
        form_data["trq_id"] = form.get("trq_id", trq_id)
        form_data["employee_id"] = form.get("employee_id", "").strip()
        form_data["sample_return"] = form.get("sample_return", "")
        form_data["remark"] = form.get("remark", "").strip()
        form_data["finishing_type"] = finishing_type
        form_data["material_type"] = material_type

        # giữ lại edit_idx qua POST nếu có
        if edit_idx is not None:
            form_data["edit_idx"] = edit_idx

        def na_or_value(key):
            na_key = key + "_na"
            if form.get(na_key):
                return "N/A"
            return form.get(key, "").strip()

        if test_group == "FINISHING TEST" and not finishing_type:
            missing_fields.append("finishing_type")
            error = "Phải chọn QA TEST hoặc LINE TEST!"

        # --- Rule riêng: Nếu Department = VFR5 thì Subcon bắt buộc và không được N/A ---
        department = form.get("department", "").strip()
        subcon_val = form.get("subcon", "").strip()
        subcon_na  = form.get("subcon_na")

        if department.upper() == "VFR5":
            if not subcon_val or subcon_na:
                missing_fields.append("subcon")
                error = "If Department is VFR5, you need to fill Subcon."

        # Nếu có thiếu, trả về form kèm lỗi
        if missing_fields:
            if not error:
                error = "Vui lòng điền đủ các trường bắt buộc (*)"
            return render_template(
                "tfr_request_form.html",
                darkmode=session.get("darkmode", "0"),
                lang=session.get("lang", "vi"),
                error=error,
                form_data=form_data,
                missing_fields=missing_fields,
                editing=bool(edit_idx is not None),
                trq_id=trq_id,
                edit_idx=edit_idx
            )

        # ---- Build new_request từ form ----
        # Lấy request_date: nếu user để trống -> dùng rule 15:00 (prefill vẫn cho sửa)
        request_date_input = (form.get("request_date") or "").strip()
        if not request_date_input:
            request_date_input = compute_request_date_now()

        item_code = na_or_value("item_code")
        supplier = na_or_value("supplier")
        subcon = na_or_value("subcon")

        test_status = form.get("test_status")
        if test_status == "nth":
            nth = form.get("test_status_nth", "").strip()
            test_status = nth + "th" if nth.isdigit() else "nth"

        remark = form.get("remark", "").strip()
        if test_group == "FINISHING TEST" and finishing_type:
            remark = f"{remark} ({finishing_type})" if remark else finishing_type

        new_request = {
            "trq_id": form.get("trq_id", trq_id),
            "requestor": form.get("requestor"),
            "employee_id": form.get("employee_id", ""),
            "department": form.get("department"),
            "request_date": request_date_input,  # <-- cho sửa, nhưng nếu trống đã auto set ở trên
            "sample_description": na_or_value("sample_description"),
            "item_code": item_code,
            "supplier": supplier,
            "subcon": subcon,
            "test_status": test_status,
            "furniture_testing": furniture_testing,
            "quantity": form.get("quantity"),
            "sample_return": form.get("sample_return", ""),
            "remark": remark,
            "test_group": test_group,
            "finishing_type": finishing_type,
            "material_type": material_type,
            "status": "Submitted",
            "decline_reason": "",
            "report_no": ""
        }

        # ✅ Tự tính ETD theo rule + tải, dựa trên danh sách hiện có (để đếm theo request_date & group)
        new_request["etd"] = calculate_default_etd(
            new_request.get("request_date", ""),
            new_request.get("test_group", ""),
            all_reqs=tfr_requests   # <— thêm dòng này
        )

        # Nếu là EDIT: giữ lại các trường hệ thống cũ (PDF/DOCX/report_no/etd/status/decline_reason)
        if editing or (trq_id and edit_idx is not None):
            try:
                _edit_idx_int = int(edit_idx)
                matches = [i for i, req in enumerate(tfr_requests) if req.get("trq_id") == trq_id]
                if len(matches) > _edit_idx_int:
                    old = tfr_requests[matches[_edit_idx_int]]
                    for k in ("status", "report_no", "pdf_path", "docx_path", "etd", "decline_reason"):
                        if k in old and old.get(k) not in (None, ""):
                            new_request[k] = old.get(k)
            except Exception:
                pass

        # ---- Ghi đè item cũ hoặc append mới ----
        if trq_id and edit_idx is not None:
            try:
                _abs = int(edit_idx)
                if 0 <= _abs < len(tfr_requests) and tfr_requests[_abs].get("trq_id") == trq_id:
                    tfr_requests[_abs] = new_request
                else:
                    # Fallback theo ordinal trong nhóm cùng trq_id
                    matches = [i for i, req in enumerate(tfr_requests) if req.get("trq_id") == trq_id]
                    if len(matches) > _abs:
                        tfr_requests[matches[_abs]] = new_request
                    else:
                        tfr_requests.append(new_request)
            except Exception:
                tfr_requests.append(new_request)
        else:
            # Tạo mới: nếu chưa có TRQ-ID (ví dụ truy cập trực tiếp POST), thì sinh mới
            if not new_request.get("trq_id"):
                existing_ids = {r.get("trq_id") for r in tfr_requests if r.get("trq_id")}
                new_request["trq_id"] = generate_unique_trq_id(existing_ids)
            tfr_requests.append(new_request)

        # ẢNH BAN ĐẦU (INITIAL PRODUCT IMAGE)
        init_files = request.files.getlist("initial_img")  # input name="initial_img" + multiple
        delete_flag = (form.get("delete_initial_img") == "1")

        # Lấy ảnh cũ nếu đang edit (để giữ nguyên khi không upload mới)
        old_initial_img = None
        old_initial_images = []
        if editing:
            old_list = safe_read_json(TFR_LOG_FILE) or []
            try:
                idx_keep = int(form.get("edit_idx", "-1"))
                if 0 <= idx_keep < len(old_list):
                    old_initial_img = old_list[idx_keep].get("initial_img")
                    old_initial_images = old_list[idx_keep].get("initial_images") or []
                    # nếu bản cũ chỉ có initial_img (chuỗi), convert thành list cho đồng bộ
                    if (not old_initial_images) and isinstance(old_initial_img, str) and old_initial_img:
                        old_initial_images = [old_initial_img]
            except Exception:
                pass

        new_initial_images = []

        if delete_flag:
            # Người dùng yêu cầu xóa ảnh initial khi edit
            new_request["initial_img"] = None
            new_request["initial_images"] = []
        else:
            if init_files:
                # Có upload mới: lưu tối đa 2 ảnh hợp lệ
                for f in init_files[:2]:
                    if not f or not f.filename:
                        continue
                    saved = _save_initial_img(f, new_request["trq_id"])  # trả về "TFR_INIT/xxx.ext"
                    if saved:
                        new_initial_images.append(saved)

                if new_initial_images:
                    new_request["initial_images"] = new_initial_images
                    new_request["initial_img"] = new_initial_images[0]  # giữ key cũ cho UI cũ
                else:
                    # Không có file hợp lệ -> nếu đang edit thì giữ ảnh cũ, ngược lại None
                    if editing and old_initial_images:
                        new_request["initial_images"] = old_initial_images
                        new_request["initial_img"] = old_initial_images[0]
                    else:
                        new_request["initial_images"] = []
                        new_request["initial_img"] = None
            else:
                # Không upload mới -> nếu edit thì giữ ảnh cũ, nếu tạo mới thì None
                if editing and old_initial_images:
                    new_request["initial_images"] = old_initial_images
                    new_request["initial_img"] = old_initial_images[0]
                else:
                    new_request["initial_images"] = []
                    new_request["initial_img"] = None

        # Ghi log như cũ
        safe_write_json(TFR_LOG_FILE, tfr_requests)
        safe_append_backup_json(TFR_LOG_FILE, new_request)

        message = (
            f"📝 [TRF] Có yêu cầu Test Request mới!\n"
            f"- Người gửi: {new_request.get('requestor')}\n"
            f"- Bộ phận: {new_request.get('department')}\n"
            f"- Ngày gửi: {new_request.get('request_date')}\n"
            f"- Nhóm test: {new_request.get('test_group')}\n"
            f"- Số lượng: {new_request.get('quantity')}\n"
            f"- Mã TRQ-ID: {new_request.get('trq_id')}"
        )
        send_teams_message(TEAMS_WEBHOOK_URL_TRF, message)

        return redirect(url_for('tfr_request_status'))

    # ===== GET lần đầu (không EDIT) -> auto fill employee_id, requestor từ session
    if not editing:
        staff_id_full = session.get("staff_id", "").strip()
        if staff_id_full and "-" in staff_id_full:
            emp_id, name = staff_id_full.split("-", 1)
            emp_id = emp_id.strip()
            name = name.strip()
        else:
            emp_id = staff_id_full
            name = ""
        form_data.setdefault("employee_id", emp_id)
        form_data.setdefault("requestor", name)

    # Tạo TRQ-ID mới nếu chưa có
    if not form_data.get("trq_id"):
        form_data["trq_id"] = generate_unique_trq_id({r.get("trq_id") for r in tfr_requests if "trq_id" in r})

    # Prefill request_date theo rule 15:00 (nhưng user vẫn có thể sửa ở form)
    if not form_data.get("request_date"):
        form_data["request_date"] = compute_request_date_now()

    return render_template(
        "tfr_request_form.html",
        darkmode=session.get("darkmode", "0"),
        lang=session.get("lang", "vi"),
        error=error,
        form_data=form_data,
        missing_fields=missing_fields,
        editing=editing,
        trq_id=trq_id,
        edit_idx=edit_idx
    )

PENDING_LOCK = Lock()
CANCEL_FLAGS = {} 

def _read_pending():
    return safe_read_json(TFR_LOG_FILE)

def _write_pending(new_list):
    safe_write_json(TFR_LOG_FILE, new_list)

def _merge_update_etd(updates):
    """
    Cập nhật ETD an toàn theo dữ liệu mới nhất trong file:
    - Nếu update có cả idx & trq_id: ưu tiên khớp trq_id, rồi mới rơi về idx.
    - Nếu chỉ có trq_id: dùng trq_id.
    - Nếu chỉ có idx: dùng idx, nhưng vẫn check bounds.
    """
    with PENDING_LOCK:
        cur = _read_pending()
        # Tạo map {trq_id: index} trên dữ liệu MỚI NHẤT
        id_to_idx = {}
        for i, r in enumerate(cur):
            tid = (r.get("trq_id") or "").strip()
            if tid:
                id_to_idx[tid] = i

        changed = False
        for u in updates:
            tid = (u.get("trq_id") or "").strip()
            etd = (u.get("etd") or "").strip()
            idx = u.get("idx")

            # Ưu tiên dùng trq_id
            if tid and tid in id_to_idx:
                cur[id_to_idx[tid]]["etd"] = etd
                changed = True
            # Fallback dùng idx nếu hợp lệ
            elif isinstance(idx, int) and 0 <= idx < len(cur):
                cur[idx]["etd"] = etd
                changed = True

        if changed:
            _write_pending(cur)
        return cur  # trả về snapshot mới nhất sau khi đã merge ETD

def _remove_approved_from_file(approved_trq_ids):
    """
    Xóa các request đã Approved RA KHỎI FILE theo trq_id (merge an toàn):
    - Luôn đọc file mới nhất
    - Lọc bỏ các phần tử có trq_id thuộc tập approved_trq_ids
    - Không đụng chạm các request mới phát sinh
    """
    if not approved_trq_ids:
        return

    with PENDING_LOCK:
        cur = _read_pending()
        keep = []
        approved_set = {tid.strip() for tid in approved_trq_ids if tid}
        for r in cur:
            tid = (r.get("trq_id") or "").strip()
            if tid and tid in approved_set:
                continue  # bỏ các request vừa approve
            keep.append(r)
        _write_pending(keep)

def make_id_index_map(pending_list):
    """
    (giữ nếu bạn đang gọi nơi khác) – map {trq_id: last_index}
    """
    mapping = {}
    if not isinstance(pending_list, list):
        return mapping
    for i, row in enumerate(pending_list):
        try:
            tid = (row.get("trq_id") or "").strip()
        except Exception:
            tid = ""
        if tid:
            mapping[tid] = i
    return mapping

# --- HÀM DUYỆT 1 REQUEST (giữ nguyên nếu app bạn đang xài) ---
def approve_all_one(req):
    """
    Approve 1 request:
      - cấp report_no + tạo DOCX/PDF
      - cập nhật Excel + TRF.xlsx
      - đẩy vào archive
      - trả về req đã cập nhật (status/report_no/pdf_path/docx_path)
    """
    # 🔒 Khoá file-level cho cả phiên cấp số -> ghi Excel -> archive
    with report_lock():                              # ← THÊM DÒNG NÀY
        # (giữ thread-lock như cũ để an toàn trong cùng process)
        with REPORT_NO_LOCK:
            current_list = safe_read_json(TFR_LOG_FILE)
            pdf_path, report_no = allocate_unique_report_no(
                approve_request_fill_docx_pdf, req, current_list
            )

        req["status"] = "Approved"
        req["decline_reason"] = ""
        req["report_no"] = report_no

        output_folder = os.path.join('static', 'TFR')
        output_docx = os.path.join(output_folder, f"{report_no}.docx")
        output_pdf  = os.path.join(output_folder, f"{report_no}.pdf")

        try:
            if not os.path.exists(output_pdf):
                try_convert_to_pdf(output_docx, output_pdf)
        except Exception as _pdf_e:
            print("PDF convert failed, fallback to DOCX:", _pdf_e)

        if os.path.exists(output_pdf):
            req['pdf_path'] = f"TFR/{report_no}.pdf"
            req['docx_path'] = None
        else:
            req['pdf_path'] = None
            req['docx_path'] = f"TFR/{report_no}.docx"

    # Ghi Excel & TRF.xlsx & archive (giữ nguyên như bạn đang có)
        try:
            write_tfr_to_excel(local_main, report_no, req)
            wb = load_workbook(local_main)
            ws = wb.active
            report_col = get_col_idx(ws, "report#")
            row_idx = None
            for row in range(2, ws.max_row + 1):
                v = ws.cell(row=row, column=report_col).value
                if v and str(v).strip() == str(report_no):
                    row_idx = row
                    break
            if row_idx:
                def set_val(col_name, value, is_date_col=False):
                    col_idx = get_col_idx(ws, col_name)
                    if col_idx:
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if is_date_col:
                            dt_val = try_parse_excel_date(value)
                            if dt_val:
                                cell.value = dt_val
                                cell.number_format = 'dd-mmm'   # <- đổi d-mmm -> dd-mmm
                            else:
                                cell.value = value
                        else:
                            cell.value = value.upper() if isinstance(value, str) else value

                def clean_type_of(val):
                    return val[:-5].strip() if val and isinstance(val, str) and val.upper().endswith(" TEST") else val

                set_val("item#", req.get("item_code", ""))
                set_val("type of", clean_type_of(req.get("test_group", "")))
                set_val("item name/ description", req.get("sample_description", ""))
                set_val("furniture testing", req.get("furniture_testing", ""))
                set_val("submiter in", req.get("requestor", ""))
                set_val("submited", req.get("department", ""))
                set_val("qa comment", req.get("remark", ""))

                etd_val = req.get("etd", "")
                set_val("etd", etd_val, is_date_col=True)  # <-- bỏ format_excel_date_short


                vn_tz = pytz.timezone("Asia/Ho_Chi_Minh")
                req_login = (req.get("request_date") or "").strip()
                val_for_excel = req_login if req_login else datetime.now(vn_tz).strftime("%Y-%m-%d")
                set_val("log in date", val_for_excel, is_date_col=True)

                finishing_type = req.get("finishing_type", "")
                material_type  = req.get("material_type", "")
                cat_comp_pos   = get_category_component_position(finishing_type, material_type)
                set_val("category / component name / position", cat_comp_pos)
                wb.save(local_main)
        except Exception as e:
            print("Ghi vào Excel bị lỗi:", e)

        try:
            append_row_to_trf(report_no, local_main, "TRF.xlsx", trq_id=req.get("trq_id", ""))
        except Exception as e:
            print("Append TRF lỗi:", e)

        try:
            vn_tz = pytz.timezone("Asia/Ho_Chi_Minh")
            short_data = {
                "trq_id": req.get("trq_id", ""),
                "report_no": req.get("report_no", ""),
                "requestor": req.get("requestor", ""),
                "department": req.get("department", ""),
                "request_date": req.get("request_date", ""),
                "item_code": req.get("item_code", ""),
                "status": req.get("status", ""),
                "pdf_path": req.get("pdf_path"),
                "docx_path": req.get("docx_path"),
                "employee_id": req.get("employee_id", ""),
                "approved_date": datetime.now(vn_tz).strftime("%Y-%m-%d"),
                "test_group": req.get("test_group", ""),
            }
            archive_request(short_data)
        except Exception as e:
            print("Archive lỗi:", e)

    return req

# ================== ROUTE: APPROVE ALL (STREAM) — ĐÃ SỬA ==================
@app.post("/approve_all_stream")
def approve_all_stream():
    """
    Sửa chính:
      1) Cập nhật ETD theo file MỚI NHẤT (merge) => không đè mất request mới.
      2) Sau MỖI request được approve, xóa request đó khỏi file bằng phép "lọc theo trq_id"
         trên dữ liệu MỚI NHẤT => không bao giờ overwrite các request mới vừa được gửi.
      3) Không còn final write "ghi đè cả file" theo snapshot cũ nữa.
    """
    def gen():
        run_id = str(uuid4())
        CANCEL_FLAGS[run_id] = False

        # Nhận input
        try:
            data = request.get_json(silent=True) or {}
            updates = data.get("updates", []) or []
        except Exception as e:
            yield json.dumps({"type": "error", "message": f"Parse JSON: {e}"}) + "\n"
            CANCEL_FLAGS.pop(run_id, None)
            return

        # (1) Merge cập nhật ETD vào file hiện tại (an toàn)
        try:
            pending_after_etd = _merge_update_etd(updates)
        except Exception as e:
            yield json.dumps({"type": "error", "message": f"Bulk ETD update: {e}"}) + "\n"
            pending_after_etd = _read_pending()

        # (2) Lập danh sách cần duyệt (submitted + có ETD)
        id_to_idx = make_id_index_map(pending_after_etd)
        todo = []
        for u in updates:
            idx = u.get("idx")
            tid = (u.get("trq_id") or "").strip()

            # Ưu tiên idx nếu còn hợp lệ và khớp trq_id (nếu có)
            picked = None
            if isinstance(idx, int) and 0 <= idx < len(pending_after_etd):
                item = pending_after_etd[idx]
                if item and item.get("status") == "Submitted" and (item.get("etd") or "").strip():
                    if not tid or tid == (item.get("trq_id") or "").strip():
                        picked = (idx, (item.get("trq_id") or "").strip(), item)

            # Fallback theo trq_id
            if not picked and tid and tid in id_to_idx:
                j = id_to_idx[tid]
                item = pending_after_etd[j]
                if item and item.get("status") == "Submitted" and (item.get("etd") or "").strip():
                    picked = (j, tid, item)

            if picked:
                todo.append(picked)
        
        def _parse_dt(s: str):
            s = (s or "").strip()
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    pass
            return datetime.max

        def _norm_type(rec: dict):
            t = (rec.get("type_of_test") or rec.get("test_group") or "")
            return t.replace(" TEST", "").strip().lower()

        # todo là list (idx, trq_id, item)
        todo.sort(key=lambda x: (
            _parse_dt(x[2].get("request_date")),
            _norm_type(x[2]),
            (x[2].get("trq_id") or "")
        ))

        yield json.dumps({"type": "start", "total": len(todo), "run_id": run_id}) + "\n"

        # (3) Duyệt từng request + mỗi lần xong thì gỡ khỏi file bằng merge-remove
        done = 0
        approved_tids = []

        for _, tid, item in todo:
            try:
                approved = approve_all_one(dict(item))  # dùng bản copy để tránh side-effect
                report_no = (approved or {}).get("report_no") or item.get("report_no")

                # Ghi nhận tiến độ
                done += 1
                approved_tids.append(tid)
                yield json.dumps({
                    "type": "progress",
                    "done": done,
                    "total": len(todo),
                    "trq_id": tid,
                    "report_no": report_no
                }) + "\n"

                # Xóa request đã approve ra khỏi file (MERGE theo trạng thái file mới nhất)
                _remove_approved_from_file([tid])

                # Người dùng bấm Cancel -> dừng sau khi xong request hiện tại
                if CANCEL_FLAGS.get(run_id):
                    yield json.dumps({"type": "cancelled", "done": done, "total": len(todo)}) + "\n"
                    CANCEL_FLAGS.pop(run_id, None)
                    return

            except Exception as e:
                yield json.dumps({"type": "error", "message": str(e), "trq_id": tid}) + "\n"

        # (4) Kết thúc bình thường
        yield json.dumps({"type": "done", "done": done, "total": len(todo)}) + "\n"
        CANCEL_FLAGS.pop(run_id, None)

    return Response(stream_with_context(gen()), mimetype="application/json")


# (tuỳ chọn) Route cancel giữ nguyên
@app.post("/approve_all_cancel")
def approve_all_cancel():
    data = request.get_json(silent=True) or {}
    run_id = data.get("run_id")
    if not run_id:
        return jsonify(success=False, message="Thiếu run_id"), 400
    if run_id not in CANCEL_FLAGS:
        return jsonify(success=False, message="Run ID không tồn tại hoặc đã kết thúc"), 404
    CANCEL_FLAGS[run_id] = True
    return jsonify(success=True)

@app.route("/tfr_request_status", methods=["GET", "POST"])
def tfr_request_status():
    try:
        import pandas as pd
    except Exception:
        pd = None

    # ===== Helpers nhỏ trong route =====
    def _parse_date(s):
        if not s:
            return datetime.max
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                pass
        return datetime.max

    def _norm_type(r):
        t = (r.get("type_of_test") or r.get("test_group") or "").strip()
        return t.replace(" TEST", "").strip().lower()

    def _tie_break(r):
        tid = (r.get("trq_id") or "").strip()
        m = re.search(r"(\d+)$", tid)
        return int(m.group(1)) if m else tid

    def _redirect_back():
        back = request.form.get("return_url")
        if back:
            return redirect(back)
        return redirect(url_for('tfr_request_status'))

    # ===== Load & quyền =====
    tfr_requests = safe_read_json(TFR_LOG_FILE) or []
    is_admin = session.get("user_type") in ("stl", "superadmin")

    # ===== Lấy Staff ID & tách (ID - Tên) =====
    viewer_staff_id = (session.get("staff_id") or request.args.get("staff_id") or "").strip()
    if viewer_staff_id and "-" in viewer_staff_id:
        _emp_id, _name = viewer_staff_id.split("-", 1)
        viewer_emp_id = _emp_id.strip()
        viewer_name   = _name.strip()
    else:
        viewer_emp_id = ""
        viewer_name   = viewer_staff_id.strip()

    def _eq(a, b):
        return (str(a or "").strip().lower() == str(b or "").strip().lower())

    # ===== Lọc hiển thị: user thường chỉ thấy request của mình (Tên HOẶC Employee ID) =====
    if not is_admin and (viewer_name or viewer_emp_id):
        tfr_requests = [
            r for r in tfr_requests
            if _eq(r.get("requestor"), viewer_name) or _eq(r.get("employee_id"), viewer_emp_id)
        ]

    # ===== POST actions =====
    if request.method == "POST":
        action = request.form.get("action")
        current = safe_read_json(TFR_LOG_FILE) or []   # snapshot mới nhất

        # ---------- APPROVE ALL ----------
        if is_admin and action == "approve_all":
            approved_count = 0
            new_pending = []
            for req in current:
                if (req.get("status") == "Submitted") and (req.get("etd") or "").strip():
                    try:
                        # ✅ log_in_date = request_date
                        req["log_in_date"] = req.get("request_date")
                        approve_all_one(req)
                        approved_count += 1
                        continue
                    except Exception as e:
                        print("Approve one (approve_all) error:", e)
                new_pending.append(req)

            safe_write_json(TFR_LOG_FILE, new_pending)
            flash(f"Đã duyệt {approved_count} request!")
            return _redirect_back()

        # ---------- APPROVE SINGLE ----------
        elif is_admin and action == "approve":
            trq_id = request.form.get("trq_id")
            edit_idx = int(request.form.get("edit_idx", 0)) if "edit_idx" in request.form else None
            matches = [i for i, req in enumerate(current) if req.get("trq_id") == trq_id]
            idx = matches[edit_idx] if edit_idx is not None and edit_idx < len(matches) else (matches[0] if matches else None)
            if idx is not None:
                req = current[idx]
                etd = (request.form.get("etd", "") or "").strip()
                if not etd:
                    flash("Bạn cần điền Estimated Completion Date (ETD) trước khi approve!")
                    return _redirect_back()

                req["etd"] = etd
                req["estimated_completion_date"] = etd
                # ✅ log_in_date = request_date
                req["log_in_date"] = req.get("request_date")

                try:
                    approve_all_one(req)
                    del current[idx]
                    safe_write_json(TFR_LOG_FILE, current)
                except Exception as e:
                    print("Approve one (single) error:", e)
                    flash("Có lỗi khi approve, vui lòng thử lại.")
            return _redirect_back()

        # ---------- DECLINE ----------
        elif is_admin and action == "decline":
            trq_id = request.form.get("trq_id")
            reason = (request.form.get("decline_reason", "") or "").strip()
            matches = [i for i, req in enumerate(current) if req.get("trq_id") == trq_id]
            idx = matches[0] if matches else None
            if idx is not None:
                current[idx]["status"] = "Declined"
                current[idx]["decline_reason"] = reason
            safe_write_json(TFR_LOG_FILE, current)
            return _redirect_back()

        # ---------- DUPLICATE ----------
        elif action == "duplicate":
            trq_id  = request.form.get("trq_id")
            edit_idx = int(request.form.get("edit_idx", 0)) if "edit_idx" in request.form else None

            matches = [i for i, req in enumerate(current) if str(req.get("trq_id")) == str(trq_id)]
            idx = matches[edit_idx] if (edit_idx is not None and 0 <= edit_idx < len(matches)) else (matches[0] if matches else None)

            if idx is not None:
                old_req = current[idx]
                new_req = old_req.copy()

                # reset fields cho bản dup
                new_req["report_no"] = ""
                new_req["status"] = "Submitted"
                new_req["pdf_path"] = ""
                new_req["decline_reason"] = ""

                if is_admin:
                    # Admin: giữ nguyên TRQ-ID (hành vi cũ)
                    # -> vẫn chèn ngay sau bản gốc để tiện edit
                    insert_pos = idx + 1
                    current.insert(insert_pos, new_req)
                    safe_write_json(TFR_LOG_FILE, current)
                    # Admin vẫn quay về trang danh sách như cũ
                    return _redirect_back()
                else:
                    # Xác thực chủ sở hữu theo TÊN hoặc EMPLOYEE ID
                    viewer_staff_id_post = (session.get("staff_id") or request.form.get("staff_id") or request.args.get("staff_id") or "").strip()
                    if viewer_staff_id_post and "-" in viewer_staff_id_post:
                        _emp2, _name2 = viewer_staff_id_post.split("-", 1)
                        owner_emp_id = _emp2.strip()
                        owner_name   = _name2.strip()
                    else:
                        owner_emp_id = ""
                        owner_name   = viewer_staff_id_post.strip()

                    is_owner = _eq(old_req.get("requestor"), owner_name) or _eq(old_req.get("employee_id"), owner_emp_id)
                    if not is_owner:
                        return _redirect_back()

                    # Người thường: tạo TRQ mới + request_date & ETD mới (luôn tính ETD)
                    existing_ids = [str(r.get("trq_id")) for r in current if r.get("trq_id")]
                    new_req["trq_id"] = generate_unique_trq_id(existing_ids)
                    new_req["request_date"] = compute_request_date_now()
                    new_req["etd"] = calculate_default_etd(
                        new_req["request_date"],
                        new_req.get("test_group", ""),
                        all_reqs=current
                    )
                    new_req["estimated_completion_date"] = new_req["etd"]

                    # Chèn ngay sau bản gốc
                    insert_pos = idx + 1
                    current.insert(insert_pos, new_req)
                    safe_write_json(TFR_LOG_FILE, current)

                    # 🔁 NEW: Sau khi Dup thành công, chuyển thẳng tới form edit của bản mới
                    return redirect(url_for(
                        'tfr_request_form',
                        trq_id=new_req["trq_id"],
                        edit_idx=insert_pos
                    ))

            return _redirect_back()

        # ---------- DELETE ----------
        elif action == "delete":
            trq_id = request.form.get("trq_id")
            edit_idx = request.form.get("edit_idx")
            if edit_idx is not None:
                try:
                    edit_idx = int(edit_idx)
                    deleted_req = current.pop(edit_idx)
                    send_teams_message(
                        TEAMS_WEBHOOK_URL_TRF,
                        f"🗑️ [TRF] Đã có yêu cầu bị xóa!\n- TRQ-ID: {deleted_req.get('trq_id')}\n- Người thao tác: {session.get('staff_id', 'Không rõ')}"
                    )
                except Exception as e:
                    print("Xóa bị lỗi:", e)
            else:
                for i, req in enumerate(current):
                    if req.get("trq_id") == trq_id:
                        deleted_req = current.pop(i)
                        send_teams_message(
                            TEAMS_WEBHOOK_URL_TRF,
                            f"🗑️ [TRF] Đã có yêu cầu bị xóa!\n- TRQ-ID: {deleted_req.get('trq_id')}\n- Người thao tác: {session.get('staff_id', 'Không rõ')}"
                        )
                        break
            safe_write_json(TFR_LOG_FILE, current)
            return _redirect_back()

    # ===== GET view (KHÔNG reload lại full list; dùng danh sách đã lọc) =====
    sort_mode = request.args.get("sort", "date")

    pairs_declined  = [(i, r) for i, r in enumerate(tfr_requests) if (r.get("status") or "").strip() == "Declined"]
    pairs_submitted = [(i, r) for i, r in enumerate(tfr_requests) if (r.get("status") or "").strip() == "Submitted"]

    if sort_mode == "type":
        key_fn = lambda it: (_norm_type(it[1]), _parse_date(it[1].get("request_date")), _tie_break(it[1]))
        pairs_declined.sort(key=key_fn)
        pairs_submitted.sort(key=key_fn)
        ordered_pairs = pairs_declined + pairs_submitted
    else:
        ordered_pairs = pairs_declined + pairs_submitted  # giữ thứ tự JSON

    real_indices  = [i for i, _ in ordered_pairs]
    show_requests = [r.copy() for _, r in ordered_pairs]  # copy để gán _rank

    # ===== Tính thứ tự trong ngày theo nhóm (để tô màu) =====
    if is_admin:
        # 1) Seed bộ đếm từ archive theo (date, group) -> count TRQ DUY NHẤT
        try:
            archive_all = safe_read_json(ARCHIVE_LOG) or []
        except Exception:
            archive_all = []

        rep2grp = _build_reportno_to_group_map()

        # 1) Seed theo TRQ duy nhất đã có trong archive (Approved) theo (ngày, nhóm)
        base_seen = {}   # (date, group) -> set(TRQ)
        for a in archive_all:
            try:
                d0  = (a.get("request_date") or "").strip()
                rep = (a.get("report_no") or "").strip()
                g0  = rep2grp.get(rep, "")
                tid = (a.get("trq_id") or "").strip() or rep  # fallback report_no
                if not (d0 and g0 and tid):
                    continue
                key = (d0, _group_of(g0))
                s = base_seen.get(key)
                if s is None:
                    s = set()
                    base_seen[key] = s
                s.add(tid)
            except Exception:
                continue

        base_count = {k: len(v) for k, v in base_seen.items()}  # (date, group) -> số TRQ duy nhất đã có

        # 2) Duyệt các request đang hiển thị và đánh số tiếp THEO TRQ DUY NHẤT
        running_seen = {}   # (date, group) -> set(TRQ) đã gặp trong batch hiện tại
        running_rank = {}   # (date, group) -> {trq: rank}
        running_count = {}  # (date, group) -> next rank (khởi từ base_count)

        for r in show_requests:
            d  = (r.get("request_date") or "").strip()
            g  = _group_of(r.get("test_group") or r.get("type_of_test"))
            st = (r.get("status") or "").strip()
            key = (d, g)

            # chuẩn bị cấu trúc
            if key not in running_seen:
                running_seen[key] = set()
                running_rank[key] = {}
                running_count[key] = base_count.get(key, 0)

            # mặc định
            r["_rank_color"] = None
            r["_group_norm"] = g

            if st != "Submitted":
                # chỉ tô màu cho Submitted theo yêu cầu
                continue

            trq = (r.get("trq_id") or "").strip()
            if not trq:
                # tránh vỡ: nếu thiếu TRQ thì coi mỗi dòng 1 "TRQ"
                trq = f"__row_{id(r)}"

            if trq in running_rank[key]:
                # dòng thứ 2, thứ 3... của cùng TRQ -> dùng lại cùng rank
                r["_rank_color"] = running_rank[key][trq]
            else:
                # lần đầu gặp TRQ này trong batch
                # Nếu TRQ đã nằm trong seed (Approved cùng ngày, nhóm), rank lịch sử của nó <= base_count
                if trq in base_seen.get(key, set()):
                    # set rank = base_count hiện tại (đủ để xác định qua/vượt mốc 5 hay chưa)
                    rank = base_count.get(key, 0)
                    running_rank[key][trq] = rank
                    running_seen[key].add(trq)
                    r["_rank_color"] = rank
                else:
                    # TRQ mới trong ngày+nhóm -> +1 theo TRQ (không theo số dòng)
                    running_count[key] += 1
                    rank = running_count[key]
                    running_rank[key][trq] = rank
                    running_seen[key].add(trq)
                    r["_rank_color"] = rank

    # ====== NEW: Banner cảnh báo từ DS: biến trong config là `local_main` ======
    def _find_local_main_path():
        candidates = []
        try:
            if local_main:
                candidates.append(local_main)
        except Exception:
            pass

        # Fallback tên mặc định
        candidates += [
            "local_main.xlsx",
            "local_main.xlsm",
            os.path.join("data", "local_main.xlsx"),
            os.path.join("data", "local_main.xlsm"),
        ]
        for p in candidates:
            try:
                if p and os.path.exists(p):
                    return p
            except Exception:
                continue
        return None  # không tìm thấy

    REPORT_RE = re.compile(r"^\d{2}-\d{1,5}$")  # YY-serial (serial 1–5 chữ số)

    def _row_is_empty_except(df_row, ignore_cols=("QR Code", "Report #")):
        """Rỗng khi tất cả cột (trừ ignore) đều trống/NaN/chuỗi trắng."""
        for col in df_row.index:
            if col in ignore_cols:
                continue
            v = df_row[col]
            if pd.isna(v):
                continue
            if isinstance(v, str):
                if v.strip() == "":
                    continue
                return False
            # số/kiểu khác -> coi là có thông tin
            return False
        return True

    def _extract_report_code(s):
        """Chuẩn hóa chuỗi thành dạng YY-serial nếu có thể."""
        if s is None:
            return None
        s = str(s).strip()
        if not s:
            return None
        if REPORT_RE.match(s):
            return s
        # normalize: 2025-123 -> 25-123 ; 25/0123 ; 25_12345 ...
        m = re.match(r"^(?:20)?(\d{2})[-_/ ]?(\d{1,5})$", s)
        if m:
            return f"{m.group(1)}-{int(m.group(2))}"
        return None

    def _pick_latest_from_codes(codes):
        """Ưu tiên năm hiện tại; cùng năm chọn serial lớn nhất; nếu không có, chọn (yy, serial) lớn nhất."""
        if not codes:
            return None
        parsed = []
        for c in codes:
            try:
                yy, ser = c.split("-")
                parsed.append((int(yy), int(ser), c))
            except Exception:
                continue
        if not parsed:
            return None
        cur_yy = int(datetime.now().strftime("%y"))
        pool = [x for x in parsed if x[0] == cur_yy] or parsed
        pool.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return pool[0][2]

    def _next_code_from_codes(codes):
        """Sinh mã kế tiếp trong năm hiện tại, serial = max+1 (không padding)."""
        cur_yy = int(datetime.now().strftime("%y"))
        max_ser = 0
        for c in codes or []:
            try:
                yy, ser = c.split("-")
                if int(yy) == cur_yy:
                    max_ser = max(max_ser, int(ser))
            except Exception:
                continue
        return f"{cur_yy}-{max_ser + 1 if max_ser >= 0 else 1}"

    def _build_ds_report_info():
        """Đọc `local_main`, xác định dòng cuối có thông tin (dòng sau rỗng), lấy mã report và trả chuỗi cảnh báo tiếng Anh."""
        if pd is None:
            return None  # không có pandas -> bỏ qua banner
        path = _find_local_main_path()
        if not path:
            return None
        try:
            df = pd.read_excel(path)
        except Exception:
            try:
                df = pd.read_excel(path, engine="openpyxl")
            except Exception:
                return None

        if df is None or df.empty:
            return None

        # Các tên cột khả dĩ cho Report #
        report_cols = ["Report #", "Report", "Report No", "Report no", "report_no", "report", "report_id", "Report_ID"]
        col_found = None
        for c in df.columns:
            if c in report_cols:
                col_found = c
                break

        # Thu thập toàn bộ code hợp lệ để dự phòng
        all_codes = []
        if col_found:
            for v in df[col_found].tolist():
                c = _extract_report_code(v)
                if c:
                    all_codes.append(c)

        # Tìm dòng cuối có thông tin: row i not-empty & row i+1 empty
        last_idx = None
        n = len(df)
        for i in range(0, n):
            this_empty = _row_is_empty_except(df.iloc[i])
            next_empty = True
            if i + 1 < n:
                next_empty = _row_is_empty_except(df.iloc[i+1])
            if (not this_empty) and next_empty:
                last_idx = i

        newest_code = None
        if last_idx is not None and col_found:
            newest_code = _extract_report_code(df.iloc[last_idx].get(col_found))

        if not newest_code:
            # fallback: lấy mã mới nhất từ toàn bộ DS hoặc sinh mã kế tiếp
            newest_code = _pick_latest_from_codes(all_codes) or _next_code_from_codes(all_codes)

        if newest_code:
            return f"⚠ The latest report: {newest_code}"
        return None

    ds_report_info = _build_ds_report_info()

    return render_template(
        "tfr_request_status.html",
        requests=show_requests,
        is_admin=is_admin,
        real_indices=real_indices,
        viewer_name=viewer_name,
        viewer_emp_id=viewer_emp_id,
        ds_report_info=ds_report_info,  # <== truyền cho template
    )

@app.route("/tfr_request_archive")
def tfr_request_archive():
    # 1) Đọc archive
    archive = safe_read_json(ARCHIVE_LOG) or []

    # 2) Gom ảnh từ toàn bộ requests (giống Status)
    try:
        tfr_all = safe_read_json(TFR_LOG_FILE) or []
    except Exception:
        tfr_all = []

    by_trq = {}
    for r in tfr_all:
        trq = (r.get("trq_id") or "").strip()
        if not trq:
            continue

        # Gom các key ảnh giống Status
        merged = {
            "initial_img": r.get("initial_img") or r.get("initial_image") or r.get("initial_image_url") or "",
            "initial_images": list(r.get("initial_images") or []),
            "form_image": r.get("form_image") or "",
            "form_images": list(r.get("form_images") or []),
            "uploaded_images": list(r.get("uploaded_images") or []),
            "product_images": list(r.get("product_images") or []),
            "images": list(r.get("images") or []),
        }

        # Nếu 'images' rỗng thì gộp tất cả mảng còn lại vào 'images'
        if not merged["images"]:
            tmp = []
            for k in ("initial_images", "form_images", "uploaded_images", "product_images"):
                vs = merged.get(k) or []
                if isinstance(vs, list):
                    tmp.extend(vs)
            # single-value
            for k in ("initial_img", "form_image"):
                v = merged.get(k)
                if isinstance(v, str) and v:
                    tmp.append(v)
            # khử trùng
            seen, out = set(), []
            for x in tmp:
                x = (x or "").strip()
                if x and x not in seen:
                    seen.add(x); out.append(x)
            merged["images"] = out

        by_trq[trq.upper()] = merged

    # Nhét ảnh vào từng record archive (không phá dữ liệu sẵn có)
    for rec in archive:
        trq = (rec.get("trq_id") or "").strip().upper()
        if not trq:
            continue
        imgpack = by_trq.get(trq)
        if imgpack:
            rec.setdefault("initial_img", imgpack.get("initial_img", ""))
            rec.setdefault("initial_images", imgpack.get("initial_images", []))
            rec.setdefault("form_image", imgpack.get("form_image", ""))
            rec.setdefault("form_images", imgpack.get("form_images", []))
            rec.setdefault("uploaded_images", imgpack.get("uploaded_images", []))
            rec.setdefault("product_images", imgpack.get("product_images", []))
            if not rec.get("images"):
                rec["images"] = imgpack.get("images", [])

    # 3) Fallback: nếu vẫn chưa có ảnh, quét static/TFR_INIT theo TRQ
    TFR_INIT_DIR = os.path.join('static', 'TFR_INIT')
    exts = ('.jpg', '.jpeg', '.png', '.webp', '.gif')

    def _find_init_by_trq(trq_id: str):
        if not trq_id:
            return []
        pattern = os.path.join(TFR_INIT_DIR, f"{trq_id}_*")
        out = []
        for p in glob.glob(pattern):
            if os.path.isfile(p) and os.path.splitext(p)[1].lower() in exts:
                # trả về đường dẫn tương đối dưới /static
                rel = os.path.relpath(p, start='static').replace('\\', '/')
                out.append(rel)
        out.sort()
        return out

    for rec in archive:
        if rec.get("initial_images") or rec.get("images") or rec.get("initial_img"):
            continue  # đã có ảnh từ JSON
        trq = (rec.get("trq_id") or "").strip()
        if not trq:
            continue
        picks = _find_init_by_trq(trq)
        if picks:
            rec["initial_images"] = picks
            rec["images"] = list(picks)

    # 4) Đọc Excel -> tạo rating_map, status_map, etd_map
    rating_map, status_map, etd_map = {}, {}, {}
    try:
        wb = safe_load_excel(local_main)  # dùng helper sẵn có
        ws = wb.active

        def find_col(*aliases):
            # thử alias trực tiếp
            for name in aliases:
                c = get_col_idx(ws, name)
                if c:
                    return c
            # fallback: quét header gần-đúng
            def norm(s): return re.sub(r"[^a-z0-9#]+", "", str(s).strip().lower())
            alias_norm = {norm(a) for a in aliases}
            # đoán mục tiêu
            want = "report"
            if any("status" in a for a in alias_norm): want = "status"
            elif any("rating" in a for a in alias_norm): want = "rating"
            elif any("etd" in a for a in alias_norm) or any("expect" in a for a in alias_norm):
                want = "etd"
            targets = {
                "status": {"status"},
                "rating": {"rating", "result"},
                "etd": {"etd", "expecteddate", "deliverydate", "expecteddelivery", "expectedfinish", "completeddate"},
                "report": {"report#", "reportno", "report", "reportnumber"},
            }[want]
            for col in range(1, ws.max_column + 1):
                h = ws.cell(row=1, column=col).value
                if h is None:
                    continue
                h_norm = norm(h)
                if h_norm in targets or any(t in h_norm for t in targets):
                    return col
            return None

        col_report = find_col("Report #", "Report#", "Report No", "Report", "report #", "report no")
        col_rating = find_col("Rating", "RATING", "rating", "Result", "RESULT", "result")
        col_status = find_col("Status", "STATUS", "status", "Current Status", "current status")
        col_etd    = find_col("ETD", "etd", "Delivery Date", "delivery date",
                              "Expected Date", "expected date",
                              "Expected Delivery", "expected delivery",
                              "Expected Finish", "expected finish",
                              "Completed Date", "completed date")

        if col_report:

            for r in range(2, ws.max_row + 1):
                key_raw = ws.cell(row=r, column=col_report).value
                if key_raw is None:
                    continue
                key = str(key_raw).strip()
                if not key:
                    continue

                # Rating / Result
                if col_rating:
                    vr = ws.cell(row=r, column=col_rating).value
                    vr_str = "" if vr is None else str(vr).strip()
                    if vr_str:
                        rating_map[key] = vr_str
                        rating_map[key.lstrip("0")] = vr_str  # fallback không 0 đầu

                # Status -> status_display
                if col_status:
                    vs = ws.cell(row=r, column=col_status).value
                    vs_str_orig = "" if vs is None else str(vs).strip()
                    vs_upper = vs_str_orig.upper()
                    if vs_upper in {"ACTIVE", "MUST", "DUE", "LATE"}:
                        disp = "ON PROGRESS"
                    elif vs_upper in {"COMPLETE", "DONE"}:
                        disp = "DONE"
                    else:
                        disp = vs_str_orig
                    status_map[key] = disp
                    status_map[key.lstrip("0")] = disp

                # ETD -> chuẩn hoá text
                if col_etd:
                    ev = ws.cell(row=r, column=col_etd).value
                    if isinstance(ev, (datetime, date)):
                        etd_text = ev.strftime("%Y-%m-%d")
                    else:
                        etd_text = ("" if ev is None else str(ev)).strip()
                        # thử parse nhanh vài format phổ biến -> 'YYYY-MM-DD'
                        if etd_text:
                            parsed = None
                            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                                try:
                                    parsed = datetime.strptime(etd_text, fmt).date()
                                    break
                                except Exception:
                                    pass
                            if parsed:
                                etd_text = parsed.strftime("%Y-%m-%d")
                    if etd_text:
                        etd_map[key] = etd_text
                        etd_map[key.lstrip("0")] = etd_text
    except Exception:
        rating_map, status_map, etd_map = {}, {}, {}

    # 5) Gắn rating/status/etd/employee_id vào từng record archive
    for rec in archive:
        rep = str(rec.get("report_no", "") or "").strip()
        if not rep:
            continue
        rec["rating"] = rating_map.get(rep, rec.get("rating", "") or "")
        rec["status_display"] = status_map.get(rep, rec.get("status_display", "") or "")
        if etd_map.get(rep):  # ETD ưu tiên lấy từ Excel theo yêu cầu
            rec["etd"] = etd_map[rep]
        rec.setdefault("employee_id", rec.get("employee_id", "") or "")

    # 6) Sắp xếp: Report No mới nằm trên
    def report_sort_key(rec):
        s = str(rec.get("report_no", "") or "")
        nums = re.findall(r"\d+", s)
        num = int(nums[-1]) if nums else -1
        return (num, s)

    archive.sort(key=report_sort_key, reverse=True)

    # đảm bảo mỗi record khi đưa sang template đều có item_code
    for rec in archive:
        # nếu trước đây lưu dưới tên 'item' hoặc field khác, rơi về luôn
        rec["item_code"] = rec.get("item_code") or rec.get("item") or ""

    # 7) Render
    return render_template("tfr_request_archive.html", requests=archive)

@app.post("/save_etd")
def save_etd():
    if not request.is_json:
        return jsonify(success=False, message="Invalid request"), 400
    data = request.get_json()
    trq_id = (data.get("trq_id") or "").strip()
    etd = (data.get("etd") or "").strip()
    if not trq_id or not etd:
        return jsonify(success=False, message="Thiếu trq_id hoặc etd"), 400

    try:
        current = _load_pending_locked()
        found = False
        for row in current:
            if (row.get("trq_id") or "").strip() == trq_id:
                row["etd"] = etd
                found = True
                break
        if not found:
            return jsonify(success=False, message="Không tìm thấy TRQ-ID trong pending!"), 404
        _write_pending_locked(current)
        return jsonify(success=True)
    except Exception as e:
        return jsonify(success=False, message="Lỗi: " + str(e)), 500

@app.route('/run_export_excel', methods=['POST'])
def run_export_excel():
    if session.get('role') not in ['stl', 'superadmin']:
        return jsonify({'success': False, 'message': 'Bạn không có quyền sử dụng chức năng này!'}), 403
    try:
        # === GỌI TRỰC TIẾP PYTHON CHẠY SCRIPT ===
        python_path = r"C:\VFR\lab_update_app\.venv\Scripts\python.exe"  # dùng python của venv
        script_path = r"C:\VFR\lab_update_app\upload TRF.py"
        result = subprocess.run([python_path, script_path],
                                shell=False, capture_output=True, text=True, timeout=900)
        if result.returncode == 0:
            return jsonify({'success': True, 'message': 'Đã chạy xong TRF!', 'reload': True})
        else:
            # Log thêm stderr nếu lỗi
            return jsonify({'success': False, 'message': f'Lỗi: {result.stderr}', 'reload': False})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Lỗi: {e}', 'reload': False})

    
@app.route("/go_report")
def go_report():
    report = request.args.get("report", "").strip()
    if report:
        return redirect(url_for("update", report=report))
    return redirect(url_for("home"))

# Trả ảnh tổng quan/cân nặng
@app.route('/images/<report>/<filename>')
def serve_general_img(report, filename):
    folder = os.path.join(UPLOAD_FOLDER, report)
    return send_from_directory(folder, filename)

@app.route("/delete_image/<report>/<imgfile>", methods=["POST"])
def delete_image_main(report, imgfile):
    img_path = os.path.join(UPLOAD_FOLDER, report, imgfile)
    # Thêm try-except để tránh lỗi race condition khi xóa cùng lúc
    try:
        if os.path.exists(img_path):
            os.remove(img_path)
    except Exception as e:
        print(f"Lỗi khi xóa ảnh: {img_path} - {e}")
    return redirect(url_for('update', report=report))

@app.route("/delete_test_group_image/<report>/<group>/<key>/<imgfile>", methods=["POST"])
def delete_test_group_image(report, group, key, imgfile):
    img_path = os.path.join(UPLOAD_FOLDER, report, imgfile)
    try:
        if os.path.exists(img_path):
            os.remove(img_path)
    except Exception as e:
        print(f"Lỗi khi xóa ảnh: {img_path} - {e}")
    return redirect(url_for("test_group_item_dynamic", report=report, group=group, test_key=key))

@app.route("/logout")
def logout():
    for k in ("auth_ok", "staff_id", "user_type", "role", "last_test_type", "last_test_code"):
        session.pop(k, None)
    return "<h3 style='text-align:center;margin-top:80px;'>Đã đăng xuất!<br><a href='/' style='color:#4d665c;'>Về trang chọn sản phẩm</a></h3>"

@app.route("/update", methods=["GET", "POST"])
def update():
    report = request.args.get("report")
    if not report:
        return redirect("/")

    item_id, row_idx = None, None
    lines = []
    message = ""
    is_logged_in = session.get("auth_ok", False)
    valid = False

    # ==== helpers cho comment file ====
    def _ensure_dir(p):
        os.makedirs(p, exist_ok=True)

    def _parse_kv_file(path):
        """Đọc file comment_{group}.txt -> dict {key: value} (key được chuẩn hóa lowercase)."""
        data = {}
        if not os.path.exists(path):
            return data
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                m = re.split(r"\s*[:=]\s*", line, maxsplit=1)
                if len(m) == 2:
                    k, v = m[0].strip(), m[1].strip()
                    if k:
                        data[k.lower()] = v
        return data

    def _upsert_kv_file(path, kv: dict):
        """Ghi/ghi đè các khóa trong file comment_{group}.txt, giữ nguyên các dòng khác."""
        lines = []
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                lines = f.readlines()

        # map key -> đã thay chưa
        replaced = {k: False for k in kv.keys()}
        new_lines = []
        for line in lines:
            line_stripped = line.strip()
            done = False
            for k, v in kv.items():
                if re.match(rf"^\s*{re.escape(k)}\s*[:=]", line_stripped):
                    new_lines.append(f"{k}: {v}\n")
                    replaced[k] = True
                    done = True
                    break
            if not done:
                new_lines.append(line)

        # append các key chưa có
        for k, v in kv.items():
            if not replaced[k]:
                new_lines.append(f"{k}: {v}\n")

        with open(path, "w", encoding="utf-8") as f:
            f.writelines(new_lines)

    # ==== đọc Excel, tìm dòng ====
    try:
        wb = safe_load_excel(local_main)
        ws = wb.active
        report_col = get_col_idx(ws, "report#") or get_col_idx(ws, "report")
        if report_col is None:
            return "❌ Không tìm thấy cột REPORT# hoặc REPORT trong file Excel!", 500

        for row in range(2, ws.max_row + 1):
            v = ws.cell(row=row, column=report_col).value
            if v and str(v).strip() == str(report):
                row_idx = row
                break
        if row_idx is None:
            return f"❌ Không tìm thấy mã report {report} trong file Excel!", 404

        valid = True

        # --- lấy rating_value để kiểm soát UI ---
        rating_value = ""
        rating_col_idx = get_col_idx(ws, "rating")
        if rating_col_idx:
            rv = ws.cell(row=row_idx, column=rating_col_idx).value
            rating_value = (str(rv).strip() if rv not in (None, "") else "")

        # --- build lines hiển thị ---
        if not is_logged_in:
            summary_keys = [
                ('TRQ ID', 'TRQ ID'),
                ('report#', 'REPORT#'),
                ('item#', 'ITEM#'),
                ('type of', 'TYPE OF'),
                ('furniture testing', 'FURNITURE TESTING'),
                ('remark', 'REMARK'),
                ('qa comment', 'QA COMMENT'),
                ('etd', 'ETD'),
                ('rating', 'RATING')
            ]
            for key, label in summary_keys:
                idx_col = get_col_idx(ws, key)
                value = ws.cell(row=row_idx, column=idx_col).value if idx_col else ""
                show_value = str(value).strip() if value not in ("", None) else ""
                lines.append((label, show_value))
        else:
            for col in range(1, ws.max_column):
                label = ws.cell(row=1, column=col).value
                value = ws.cell(row=row_idx, column=col).value
                if label and value not in (None, ""):
                    lines.append((str(label).upper(), str(value)))

    except Exception as e:
        print("Lỗi khi đọc file excel:", e)
        print(traceback.format_exc())
        return f"Lỗi khi xử lý file: {e}", 500

    # ==== lấy sẵn Sample Info (ưu tiên nhóm test gần nhất, fallback main) ====
    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    _ensure_dir(report_folder)

    group_pref = get_last_test_code(report, default="main")

    def _read_notes(gname):
        return _parse_kv_file(os.path.join(report_folder, f"comment_{gname}.txt"))

    sample_notes = _read_notes(group_pref)
    if not sample_notes or (not sample_notes.get("sample_weight") and not sample_notes.get("sample_size")):
        base_notes = _read_notes("main")
        if base_notes:
            for k, v in base_notes.items():
                sample_notes.setdefault(k, v)

    # --- Chuẩn hóa sample_weight: lấy số thuần để render vào <input type="number">
    def _extract_number_str(s):
        if not s:
            return ""
        s = str(s)
        m = re.search(r'[-+]?\d+(?:[.,]\d+)?', s)  # lấy số đầu tiên
        if not m:
            return ""
        return m.group(0).replace(",", ".")  # 12,3 -> 12.3

    sample_weight_val_raw = sample_notes.get("sample_weight", "")
    sample_weight_val = _extract_number_str(sample_weight_val_raw)

    # --- Parse sample_size "L x W x H" -> ba trường để bind vào input
    sample_size_val = sample_notes.get("sample_size", "")

    def _parse_size_triplet(s):
        if not s:
            return "", "", ""
        s = str(s)
        m = re.search(
            r'([0-9]+(?:[.,][0-9]+)?)\s*[x×*]\s*([0-9]+(?:[.,][0-9]+)?)\s*[x×*]\s*([0-9]+(?:[.,][0-9]+)?)',
            s, re.I
        )
        if not m:
            return "", "", ""
        a, b, c = m.group(1), m.group(2), m.group(3)
        return a.replace(",", "."), b.replace(",", "."), c.replace(",", ".")

    size_length_val, size_width_val, size_height_val = _parse_size_triplet(sample_size_val)

    # --- nếu chưa login: xử lý login form ---
    if not is_logged_in:
        if request.method == "POST" and request.form.get("action") == "login":
            password_input = request.form.get("password")
            if login(password_input):
                return redirect(url_for("update", report=report))
            else:
                message = "Sai mật khẩu!"

        return render_template(
            "info_line.html",
            lines=lines,
            message=message,
            logout=False,
            img_overview=get_img_urls(report, "overview"),
            img_weight=[],
            show_hint=True,
            show_func=False,
            report_id=report,
            test_groups=TEST_GROUPS,
            rating_value=rating_value,
            sample_weight=sample_weight_val,
            sample_size=sample_size_val,
            show_line_test_done_notice=False,  # ẩn phần notice khi chưa login
            so_gio_test=SO_GIO_TEST,
        )

    # ==== ĐÃ ĐĂNG NHẬP: XỬ LÝ POST ====
    if request.method == "POST":
        action = request.form.get("action")

        # --- Upload overview images ---
        if action == "upload_overview":
            files = request.files.getlist('overview_imgs')
            folder = os.path.join(UPLOAD_FOLDER, report)
            os.makedirs(folder, exist_ok=True)
            for i, file in enumerate(files):
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[1].lower()
                    filename = f"overview_{int(datetime.now().timestamp())}_{i}.{ext}"
                    file.save(os.path.join(folder, filename))
            return redirect(url_for("update", report=report))

        # --- Upload weight images ---
        elif action == "upload_weight":
            files = request.files.getlist('weight_imgs')
            folder = os.path.join(UPLOAD_FOLDER, report)
            os.makedirs(folder, exist_ok=True)
            for i, file in enumerate(files):
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[1].lower()
                    filename = f"weight_{int(datetime.now().timestamp())}_{i}.{ext}"
                    file.save(os.path.join(folder, filename))
            return redirect(url_for("update", report=report))

        # --- LƯU SAMPLE INFO vào comment_{group}.txt (KHÔNG ghi Excel) ---
        elif action == "save_sample_info":
            weight = request.form.get("sample_weight", "").strip()
            size_length = request.form.get("size_length", "").strip()
            size_width  = request.form.get("size_width", "").strip()
            size_height = request.form.get("size_height", "").strip()

            size_str = ""
            if size_length and size_width and size_height:
                size_str = f"{size_length} x {size_width} x {size_height}"

            group_code = get_last_test_code(report, default="main")
            comment_file = os.path.join(UPLOAD_FOLDER, str(report), f"comment_{group_code}.txt")
            os.makedirs(os.path.dirname(comment_file), exist_ok=True)

            updates = {}
            if weight:
                # Lưu nguyên số thập phân, thêm " kg" để người đọc dễ hiểu
                # (nếu không muốn 'kg' thì bỏ đi)
                updates["sample_weight"] = f"{weight} kg"
            if size_str:
                updates["sample_size"] = size_str

            if updates:
                _upsert_kv_file(comment_file, updates)

            # PRG: đảm bảo reload trang sẽ hiển thị giá trị mới trong input
            return redirect(url_for("update", report=report))
        
        # --- Đánh dấu "testing" ---
        elif valid and action == "testing":
            wb = safe_load_excel(local_main)
            ws = wb.active
            test_date_col = get_col_idx(ws, "test date")
            rating_col = get_col_idx(ws, "rating")
            vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
            now = datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M").upper()
            if test_date_col:
                ws.cell(row=row_idx, column=test_date_col).value = now
            if rating_col:
                ws.cell(row=row_idx, column=rating_col).value = "PENDING"
            safe_save_excel(wb, local_main)
            message = f"Đã ghi thời gian kiểm tra và cập nhật trạng thái PENDING cho {report}!"

        # --- Đánh dấu "test_done" ---
        elif valid and action == "test_done":
            wb = safe_load_excel(local_main)
            ws = wb.active
            complete_col = get_col_idx(ws, "complete date")
            vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
            now = datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M").upper()
            if complete_col:
                ws.cell(row=row_idx, column=complete_col).value = now
            safe_save_excel(wb, local_main)
            message = f"Đã ghi hoàn thành test cho {report}!"

        # --- Rating PASS/FAIL/DATA (ghi Excel như cũ) ---
        elif valid and action and action.startswith("rating_"):
            print("==> ĐANG XỬ LÝ RATING:", action, "CHO REPORT", report)
            value = action.replace("rating_", "").upper()

            wb = safe_load_excel(local_main)
            ws = wb.active

            rating_col = get_col_idx(ws, "rating")
            status_col = get_col_idx(ws, "status")
            if rating_col:
                ws.cell(row=row_idx, column=rating_col).value = value

            # (… giữ nguyên phần còn lại xử lý Teams / copy completed / log như code gốc …)
            # --- LẤY LOẠI TEST GẦN NHẤT ---
            group_code = get_last_test_code(report)
            group_title = get_group_title(group_code) if group_code else None

            if not group_code:
                last_test_type = get_last_test_type(report)
                if last_test_type:
                    for g_id, g_name in TEST_GROUPS:
                        if g_name == last_test_type:
                            group_code = g_id
                            group_title = g_name
                            break

            if not group_code:
                type_of_col = get_col_idx(ws, "type of")
                type_of = ws.cell(row=row_idx, column=type_of_col).value if type_of_col else ""
                # nếu có bảng map chuẩn hóa riêng thì áp dụng ở đây thay vì replace space.
                if type_of:
                    t_key = str(type_of).strip().lower()
                    mapped_code = None
                    try:
                        from config import TYPE_OF_MAP  # optional mapping: {raw_lower: group_code}
                    except Exception:
                        TYPE_OF_MAP = {}
                    if isinstance(TYPE_OF_MAP, dict):
                        mapped_code = TYPE_OF_MAP.get(t_key) or TYPE_OF_MAP.get(t_key.replace(" ", "_"))
                    if not mapped_code:
                        for g_id, g_name in TEST_GROUPS:
                            if t_key == str(g_name).strip().lower():
                                mapped_code = g_id
                                break
                    group_code = mapped_code or t_key.replace(" ", "_")
                else:
                    group_code = None
                group_title = get_group_title(group_code) or (type_of or "")

            country_col = get_col_idx(ws, "country of destination")
            furniture_testing_col = get_col_idx(ws, "furniture testing")
            country = ws.cell(row=row_idx, column=country_col).value if country_col else ""
            furniture_testing = ws.cell(row=row_idx, column=furniture_testing_col).value if furniture_testing_col else ""

            trq_col = get_col_idx(ws, "trq id")
            item_col = get_col_idx(ws, "item#")
            desc_col = get_col_idx(ws, "item name/ description")
            requestor_col = get_col_idx(ws, "submiter in") or get_col_idx(ws, "submitter in charge") or get_col_idx(ws, "requestor")

            trq = ws.cell(row=row_idx, column=trq_col).value if trq_col else ""
            item = ws.cell(row=row_idx, column=item_col).value if item_col else ""
            desc = ws.cell(row=row_idx, column=desc_col).value if desc_col else ""
            requestor = ws.cell(row=row_idx, column=requestor_col).value if requestor_col else ""

            report_url = f"{request.url_root.rstrip('/')}/update?report={report}"
            staff_id = session.get("staff_id", "Không rõ")

            teams_msg = None
            if value == "PASS":
                teams_msg = (
                    f"✅ **PASS** {group_title}\n"
                    f"- TRQ: {trq}\n"
                    f"- Report#: {report}\n"
                    f"- Item#: {item}\n"
                    f"- Description: {desc}\n"
                    f"- Group: {group_title}\n"
                    f"- Country of Destination: {country}\n"
                    f"- Furniture Testing: {furniture_testing}\n"
                    f"- Requestor: {requestor}\n"
                    f"- Nhân viên thao tác: {staff_id}\n"
                    f"Chi tiết: {report_url}"
                )
            elif value in ["FAIL", "DATA"]:
                report_folder = os.path.join(UPLOAD_FOLDER, str(report))
                status_file = os.path.join(report_folder, f"status_{group_code}.txt")
                comment_file = os.path.join(report_folder, f"comment_{group_code}.txt")
                group_titles = TEST_GROUP_TITLES.get(group_code, {})
                status_notes = load_group_notes(status_file)
                comment_notes = load_group_notes(comment_file)
                group_fails = []
                for key, title in group_titles.items():
                    status_val = status_notes.get(key)
                    if status_val == "FAIL":
                        comment_val = comment_notes.get(key, "")
                        group_fails.append(f"- {title['short']}: {comment_val if comment_val else '(Không có ghi chú)'}")
                status_text = "❌ **FAIL**" if value == "FAIL" else "⚠️ **DATA**"
                if group_fails:
                    teams_msg = (
                        f"{status_text} {group_title}\n"
                        f"- TRQ: {trq}\n"
                        f"- Report#: {report}\n"
                        f"- Item#: {item}\n"
                        f"- Description: {desc}\n"
                        f"- Group: {group_title}\n"
                        f"- Country of Destination: {country}\n"
                        f"- Furniture Testing: {furniture_testing}\n"
                        f"- Requestor: {requestor}\n"
                        f"- Nhân viên thao tác: {staff_id}\n"
                        f"- Các mục FAIL:\n"
                        + "\n".join(group_fails)
                        + f"\nChi tiết: {report_url}"
                    )
                else:
                    teams_msg = (
                        f"{status_text} {group_title}\n"
                        f"- TRQ: {trq}\n"
                        f"- Report#: {report}\n"
                        f"- Item#: {item}\n"
                        f"- Description: {desc}\n"
                        f"- Group: {group_title}\n"
                        f"- Country of Destination: {country}\n"
                        f"- Furniture Testing: {furniture_testing}\n"
                        f"- Requestor: {requestor}\n"
                        f"- Nhân viên thao tác: {staff_id}\n"
                        f"- Không có mục nào FAIL trong nhóm này."
                        + f"\nChi tiết: {report_url}"
                    )

            if teams_msg:
                send_teams_message(TEAMS_WEBHOOK_URL_RATE, teams_msg)

            if status_col:
                ws.cell(row=row_idx, column=status_col).value = "COMPLETE"
                fill_complete = PatternFill("solid", fgColor="BFBFBF")
                for col in range(2, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col).fill = fill_complete

            # Copy sang completed file (giữ nguyên logic cũ)
            if os.path.exists(local_complete):
                wb_c = safe_load_excel(local_complete)
                ws_c = wb_c.active
            else:
                wb_c = Workbook()
                ws_c = wb_c.active
                for col in range(1, ws.max_column + 1):
                    from_cell = ws.cell(row=1, column=col)
                    to_cell = ws_c.cell(row=1, column=col)
                    to_cell.value = from_cell.value
                    if from_cell.font:        to_cell.font = from_cell.font.copy()
                    if from_cell.border:      to_cell.border = from_cell.border.copy()
                    if from_cell.fill:        to_cell.fill = from_cell.fill.copy()
                    if from_cell.protection:  to_cell.protection = from_cell.protection.copy()
                    if from_cell.alignment:   to_cell.alignment = from_cell.alignment.copy()
                    to_cell.number_format = from_cell.number_format
                    col_letter = from_cell.column_letter
                    ws_c.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                ws_c.row_dimensions[1].height = ws.row_dimensions[1].height
                safe_save_excel(wb_c, local_complete)

            report_idx_in_c = get_col_idx(ws_c, "report#") or get_col_idx(ws_c, "report") or 2
            found_row = None
            for r in range(2, ws_c.max_row + 1):
                v = ws_c.cell(row=r, column=report_idx_in_c).value
                if v and str(v).strip().upper() == str(report).upper():
                    found_row = r
                    break

            if found_row:
                copy_row_with_style(ws, ws_c, row_idx, to_row=found_row)
            else:
                copy_row_with_style(ws, ws_c, row_idx)

            safe_save_excel(wb_c, local_complete)
            safe_save_excel(wb, local_main)

            # Log hoàn thành
            type_of_col = get_col_idx(ws, "type of")
            type_of = ws.cell(row=row_idx, column=type_of_col).value if type_of_col else ""
            vn_tz = pytz.timezone("Asia/Ho_Chi_Minh")
            now = datetime.now(vn_tz)
            tval = now.hour * 60 + now.minute
            office_start = 8 * 60
            office_end = 16 * 60 + 45
            ot_end = 23 * 60 + 59
            if office_start <= tval < office_end:
                ca = "office"
            elif office_end <= tval <= ot_end:
                ca = "ot"
            else:
                ca = ""
            employee_id = session.get("staff_id", "")
            log_report_complete(report, type_of, ca, employee_id)

            message = f"Đã cập nhật đánh giá: <b>{value}</b> cho {report}!"
            check_and_reset_counter()
            update_counter()

    # === Lấy loại test gần nhất (last_test_type) ===
    last_test_type = get_last_test_type(report)

    # === Kiểm tra đã đủ số giờ line test chưa ===
    elapsed = get_line_test_elapsed(report)
    show_line_test_done = elapsed is not None and elapsed >= SO_GIO_TEST

    # === Kiểm tra đã có ảnh after chưa ===
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    imgs_after = []
    after_tag = "line_after"
    if os.path.exists(folder):
        for f in sorted(os.listdir(folder)):
            if allowed_file(f) and f.startswith(after_tag):
                imgs_after.append(f"/images/{report}/{f}")
    has_after_img = len(imgs_after) > 0

    show_line_test_done_notice = show_line_test_done and not has_after_img

    return render_template(
        "info_line.html",
        lines=lines,
        message=message,
        logout=True,
        img_overview=get_img_urls(report, "overview"),
        img_weight=get_img_urls(report, "weight"),
        show_hint=False,
        show_func=True,
        report_id=report,
        test_groups=TEST_GROUPS,
        last_test_type=last_test_type,
        so_gio_test=SO_GIO_TEST,
        show_line_test_done_notice=show_line_test_done_notice,
        rating_value=rating_value,
        # Sample info
        sample_weight=sample_weight_val,
        sample_size=sample_size_val,
        size_length=size_length_val,
        size_width=size_width_val,
        size_height=size_height_val,
    )

def _has_images(report_folder: str, group: str, key: str, is_hotcold_like: bool) -> bool:
    if not os.path.exists(report_folder):
        return False
    try:
        files = os.listdir(report_folder)
    except Exception:
        return False

    if is_hotcold_like:
        # chấp nhận tên có/không kèm group sau before/after
        prefixes = (
            f"{key}_before_{group}",
            f"{key}_after_{group}",
            f"{key}_before_",
            f"{key}_after_",
        )
        return any(allowed_file(fn) and fn.startswith(prefixes) for fn in files)
    else:
        pref = f"test_{group}_{key}_"
        return any(allowed_file(fn) and fn.startswith(pref) for fn in files)

@app.route("/download_images/<report>")
def download_all_images(report):
    base_dir = os.path.join("images", report)

    if not os.path.exists(base_dir):
        return f"Không tìm thấy thư mục hình cho report {report}", 404

    # Tạo ZIP trong RAM
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(base_dir):
            for file in files:
                if file.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                    full_path = os.path.join(root, file)

                    # Đường dẫn bên trong ZIP giữ nguyên cấu trúc
                    arcname = os.path.relpath(full_path, base_dir)

                    zipf.write(full_path, arcname)

    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=f"{report}_images.zip",
        mimetype="application/zip"
    )
# >>> ADD
@app.get('/api/report/detect')
def api_report_detect():
    report = (request.args.get('report') or '').strip()
    if not report:
        return jsonify({'ok': False, 'error': 'missing report'}), 400

    # TODO: viết logic tự đoán nếu muốn (từ last_test_type, file/ảnh, v.v.)
    # Tạm thời: luôn yêu cầu người dùng chọn
    return jsonify({'ok': False})

# ==== ROUTE ====
@app.route("/api/report/create")
def api_report_create():
    """
    Create .docx report for any type using auto-mapping to TEST_GROUP_TITLES.
    Templates are placed next to app.py.
    """
    report_id = (request.args.get("report") or "").strip()
    rtype     = (request.args.get("type") or "").strip().lower()

    tpl_name = TEMPLATE_MAP.get(rtype)
    if not tpl_name:
        return jsonify({"ok": False, "error": f"Unknown type: {rtype}"}), 400

    tpl_path = os.path.join(os.path.dirname(__file__), tpl_name)
    if not os.path.exists(tpl_path):
        return jsonify({"ok": False, "error": f"Template not found: {tpl_name}"}), 404

    # Use the same Excel as for bed (adjust if you later separate per-type files)
    excel_path = os.path.join(local_main, "ds san pham test voi qr.xlsx")

    try:
        # Generic fill for ALL types with auto mapping
        generated_by = session.get('staff_id') or session.get('user_type') or "VFR User"

        bio = fill_cover_from_excel_generic(
            template_docx_path=tpl_path,
            excel_path_or_name=excel_path,
            report_id=report_id,
            template_key=rtype,
            generated_by=generated_by,   # NEW
        )
        return send_file(
            bio,
            as_attachment=True,
            download_name=f"{report_id or 'report'}_{rtype}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as ex:
        # Fallback: return raw template to not block users
        print("Autofill error:", ex)
        return send_file(
            tpl_path,
            as_attachment=True,
            download_name=f"{report_id or 'report'}_{rtype}_TEMPLATE.docx"
        )

# --- THAY THẾ HẲN hàm test_group_page ---
@app.route("/test_group/<report>/<group>", methods=["GET", "POST"])
def test_group_page(report, group):
    # Lưu context gần nhất
    set_last_test(report, group)

    group_titles = TEST_GROUP_TITLES.get(group)
    if not group_titles:
        return "Nhóm kiểm tra không tồn tại!", 404

    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(report_folder, exist_ok=True)

    # Nơi các trang hot/cold ghi vào:
    status_file  = os.path.join(report_folder, f"status_{group}.txt")
    comment_file = os.path.join(report_folder, f"comment_{group}.txt")
    all_status   = load_group_notes(status_file)    # {key -> PASS/FAIL/N/A/DATA...}
    all_comment  = load_group_notes(comment_file)   # {key -> comment string}

    test_status = {}

    for key in group_titles:
        # 1) Lấy từ file tổng (status_{group}.txt / comment_{group}.txt)
        st = all_status.get(key)
        cm = all_comment.get(key)

        # 2) Nếu chưa có, đọc fallback theo cả 2 pattern file riêng:
        #    - Mới:  status_{group}_{key}.txt / comment_{group}_{key}.txt
        #    - Cũ:   {key}_{group}_status.txt / {key}_{group}_comment.txt
        if not st:
            for st_path in [
                os.path.join(report_folder, f"status_{group}_{key}.txt"),
                os.path.join(report_folder, f"{key}_{group}_status.txt"),
            ]:
                if os.path.exists(st_path):
                    try:
                        v = (safe_read_text(st_path) or "").strip()
                        if v:
                            st = v
                            break
                    except Exception:
                        pass

        if not cm:
            for cm_path in [
                os.path.join(report_folder, f"comment_{group}_{key}.txt"),
                os.path.join(report_folder, f"{key}_{group}_comment.txt"),
            ]:
                if os.path.exists(cm_path):
                    try:
                        v = (safe_read_text(cm_path) or "").strip()
                        if v:
                            cm = v
                            break
                    except Exception:
                        pass

        # 3) Kiểm tra ảnh (đã ok cho cả hot_cold & thường)
        has_img = _has_images(report_folder, group, key, key in HOTCOLD_LIKE)

        test_status[key] = {"status": st, "comment": cm, "has_img": has_img}
    # Riêng tủ US (nếu có)
    f2057_status = {}
    if group == "tu_us":
        for fkey in F2057_TEST_TITLES:
            f2057_status[fkey] = get_group_test_status(report, group, fkey)

    return render_template(
        "test_group_menu.html",
        report=report,
        group=group,
        test_titles=group_titles,
        test_status=test_status,
        F2057_TEST_TITLES=F2057_TEST_TITLES,
        f2057_status=f2057_status
    )

@app.route('/test_group/<report>/<group>/<test_key>', methods=['GET', 'POST'])
def test_group_item_dynamic(report, group, test_key):
    # Lưu lại loại test gần nhất
    set_last_test(report, group)

    # Hot/Cold chuyển sang route riêng
    if test_key in HOTCOLD_LIKE and group in INDOOR_GROUPS:
        return redirect(url_for("hot_cold_test", report=report, group=group, test_key=test_key))

    # Kiểm tra tồn tại test key
    group_titles = TEST_GROUP_TITLES.get(group)
    if not group_titles or test_key not in group_titles:
        return "Mục kiểm tra không tồn tại!", 404
    title = group_titles[test_key]

    # Thư mục theo report
    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(report_folder, exist_ok=True)
    status_file = os.path.join(report_folder, f"status_{group}.txt")
    comment_file = os.path.join(report_folder, f"comment_{group}.txt")

    # Đặc thù nhóm TRANSIT
    is_rh_np = (group == "transit_RH_np")
    is_drop = (is_drop_test(title) if group.startswith("transit") else False) or (group == "transit_181_lt68" and test_key == "step4")
    is_impact = is_impact_test(title) if group.startswith("transit") else False
    is_rot = is_rotational_test(title) if group.startswith("transit") else False
    rh_impact_zones = RH_IMPACT_ZONES if is_rh_np and test_key == "step3" else []
    rh_vib_zones = RH_VIB_ZONES if is_rh_np and test_key == "step4" else []
    rh_second_impact_zones = RH_SECOND_IMPACT_ZONES if is_rh_np and test_key == "step5" else []
    rh_step12_zones = RH_STEP12_ZONES if is_rh_np and test_key == "step12" else []

    # ------------- AJAX IMAGE UPLOAD/DELETE (JSON RESPONSE) -------------
    if request.method == "POST" and request.headers.get("X-Requested-With") == "XMLHttpRequest":
        # Kiểm tra các vùng ảnh đặc biệt RH
        imgs = {}

        # ========== GT68 FACE ZONES (chỉ xử lý GT68 ở đây) ==========
        if group == "transit_181_gt68" and test_key == "step4":
            for idx, zone in enumerate(GT68_FACE_ZONES):
                files = request.files.getlist(f'gt68_face_img_{zone}')
                if files:
                    imgs[str(idx)] = []  # FIX: đồng bộ key "0".."5" để FE đọc data.imgs[zone]
                    for file in files:
                        if file and allowed_file(file.filename):
                            ext = file.filename.rsplit('.', 1)[-1].lower()
                            prefix = f"test_{group}_{test_key}_gt68_face_{zone}_"
                            nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                            next_num = max(nums, default=0) + 1
                            fname = f"{prefix}{next_num}.{ext}"
                            file.save(os.path.join(report_folder, fname))
                            imgs[str(idx)].append(f"/images/{report}/{fname}")

        # ========== RH Impact zones (tách ra ngoài nhánh GT68) ==========
        # FIX: các khối RH/Drop/Impact/Rot KHÔNG còn lồng trong nhánh GT68
        for zone, _ in rh_impact_zones:
            files = request.files.getlist(f'rh_impact_img_{zone}')
            if files:
                imgs.setdefault(zone, [])
                for file in files:
                    if file and allowed_file(file.filename):
                        ext = file.filename.rsplit('.', 1)[-1].lower()
                        prefix = f"test_{group}_{test_key}_{zone}_"
                        nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                        next_num = max(nums, default=0) + 1
                        fname = f"{prefix}{next_num}.{ext}"
                        file.save(os.path.join(report_folder, fname))
                        imgs[zone].append(f"/images/{report}/{fname}")

        # ========== RH Vib zones ==========
        for zone, _ in rh_vib_zones:
            files = request.files.getlist(f'rh_vib_img_{zone}')
            if files:
                imgs.setdefault(zone, [])
                for file in files:
                    if file and allowed_file(file.filename):
                        ext = file.filename.rsplit('.', 1)[-1].lower()
                        prefix = f"test_{group}_{test_key}_{zone}_"
                        nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                        next_num = max(nums, default=0) + 1
                        fname = f"{prefix}{next_num}.{ext}"
                        file.save(os.path.join(report_folder, fname))
                        imgs[zone].append(f"/images/{report}/{fname}")

        # ========== RH Second impact zones ==========
        for zone, _ in rh_second_impact_zones:
            files = request.files.getlist(f'rh_second_impact_img_{zone}')
            if files:
                imgs.setdefault(zone, [])
                for file in files:
                    if file and allowed_file(file.filename):
                        ext = file.filename.rsplit('.', 1)[-1].lower()
                        prefix = f"test_{group}_{test_key}_{zone}_"
                        nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                        next_num = max(nums, default=0) + 1
                        fname = f"{prefix}{next_num}.{ext}"
                        file.save(os.path.join(report_folder, fname))
                        imgs[zone].append(f"/images/{report}/{fname}")

        # ========== RH step12 zones ==========
        for zone, _ in rh_step12_zones:
            files = request.files.getlist(f'rh_step12_img_{zone}')
            if files:
                imgs.setdefault(zone, [])
                for file in files:
                    if file and allowed_file(file.filename):
                        ext = file.filename.rsplit('.', 1)[-1].lower()
                        prefix = f"test_{group}_{test_key}_{zone}_"
                        nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                        next_num = max(nums, default=0) + 1
                        fname = f"{prefix}{next_num}.{ext}"
                        file.save(os.path.join(report_folder, fname))
                        imgs[zone].append(f"/images/{report}/{fname}")

        # ========== DROP, IMPACT, ROTATION (tách ra ngoài nhánh GT68) ==========
        # Drop
        if is_drop:
            for idx, zone in enumerate(DROP_ZONES):
                files = request.files.getlist(f'drop_img_{zone}')
                if files:
                    imgs.setdefault(idx, [])
                    for file in files:
                        if file and allowed_file(file.filename):
                            ext = file.filename.rsplit('.', 1)[-1].lower()
                            prefix = f"test_{group}_{test_key}_drop_{zone}_"
                            nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                            next_num = max(nums, default=0) + 1
                            fname = f"{prefix}{next_num}.{ext}"
                            file.save(os.path.join(report_folder, fname))
                            imgs[idx].append(f"/images/{report}/{fname}")

        # Impact
        if is_impact:
            for idx, zone in enumerate(IMPACT_ZONES):
                files = request.files.getlist(f'impact_img_{zone}')
                if files:
                    imgs.setdefault(idx, [])
                    for file in files:
                        if file and allowed_file(file.filename):
                            ext = file.filename.rsplit('.', 1)[-1].lower()
                            prefix = f"test_{group}_{test_key}_impact_{zone}_"
                            nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                            next_num = max(nums, default=0) + 1
                            fname = f"{prefix}{next_num}.{ext}"
                            file.save(os.path.join(report_folder, fname))
                            imgs[idx].append(f"/images/{report}/{fname}")

        # Rotation
        if is_rot:
            for idx, zone in enumerate(ROT_ZONES):
                files = request.files.getlist(f'rot_img_{zone}')
                if files:
                    imgs.setdefault(idx, [])
                    for file in files:
                        if file and allowed_file(file.filename):
                            ext = file.filename.rsplit('.', 1)[-1].lower()
                            prefix = f"test_{group}_{test_key}_rotation_{zone}_"
                            nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                            next_num = max(nums, default=0) + 1
                            fname = f"{prefix}{next_num}.{ext}"
                            file.save(os.path.join(report_folder, fname))
                            imgs[idx].append(f"/images/{report}/{fname}")

        # THƯỜNG
        if request.files.getlist('test_imgs'):
            imgs['normal'] = []
            for file in request.files.getlist('test_imgs'):
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[-1].lower()
                    prefix = f"test_{group}_{test_key}_"
                    nums = [int(f[len(prefix):].split('.')[0]) for f in os.listdir(report_folder) if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()]
                    next_num = max(nums, default=0) + 1
                    fname = f"{prefix}{next_num}.{ext}"
                    file.save(os.path.join(report_folder, fname))
                    imgs['normal'].append(f"/images/{report}/{fname}")

        # Xóa ảnh AJAX
        if 'delete_img' in request.form:
            fname = request.form['delete_img']
            img_path = os.path.join(report_folder, fname)
            if os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass  # Đã bị xóa bởi thread khác
            # Trả lại danh sách ảnh còn lại
            if 'kind' in request.form and 'zone_idx' in request.form:
                kind = request.form['kind']
                idx = request.form['zone_idx']
                if kind in ['drop', 'impact', 'rot']:
                    # Lấy lại danh sách ảnh còn lại cho zone idx
                    if kind == 'drop':
                        zone = DROP_ZONES[int(idx)]
                        prefix = f"test_{group}_{test_key}_drop_{zone}_"
                    elif kind == 'impact':
                        zone = IMPACT_ZONES[int(idx)]
                        prefix = f"test_{group}_{test_key}_impact_{zone}_"
                    elif kind == 'rot':
                        zone = ROT_ZONES[int(idx)]
                        prefix = f"test_{group}_{test_key}_rotation_{zone}_"
                    imgs[int(idx)] = []
                    for f in os.listdir(report_folder):
                        if allowed_file(f) and f.startswith(prefix):
                            imgs[int(idx)].append(f"/images/{report}/{f}")
                elif kind == 'gt68_face' and group == "transit_181_gt68" and test_key == "step4":
                    idx = int(idx)
                    zone = GT68_FACE_ZONES[idx]
                    prefix = f"test_{group}_{test_key}_gt68_face_{zone}_"
                    imgs[str(idx)] = []  # FIX: trả về key "0".."5" để khớp FE
                    for f in os.listdir(report_folder):
                        if allowed_file(f) and f.startswith(prefix):
                            imgs[str(idx)].append(f"/images/{report}/{f}")
                else:
                    # RH zones
                    zone = idx
                    prefix = f"test_{group}_{test_key}_{zone}_"
                    imgs[zone] = []
                    for f in os.listdir(report_folder):
                        if allowed_file(f) and f.startswith(prefix):
                            imgs[zone].append(f"/images/{report}/{f}")
            elif 'delete_img' in request.form:
                # Ảnh thường
                imgs['normal'] = []
                for f in os.listdir(report_folder):
                    if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_"):
                        imgs['normal'].append(f"/images/{report}/{f}")

        return jsonify(imgs=imgs)

    # --- Trạng thái PASS/FAIL/N/A ---
    all_status = load_group_notes(status_file)
    status_value = all_status.get(test_key, "")

    # --- Comment ---
    comment = get_group_note_value(comment_file, test_key) 
    
    def get_comment(file_path, key):
        return get_group_note_value(file_path, key)

    # Lấy comment của mục này
    comment = get_comment(comment_file, test_key)

    # --- Xác định loại test đặc biệt ---
    is_rh_np = (group == "transit_RH_np")
    is_drop = (is_drop_test(title) if group.startswith("transit") else False) or (group == "transit_181_lt68" and test_key == "step4")
    is_impact = is_impact_test(title) if group.startswith("transit") else False
    is_rot = is_rotational_test(title) if group.startswith("transit") else False

    # --- RH Non Pallet zones ---
    rh_impact_zones = RH_IMPACT_ZONES if is_rh_np and test_key == "step3" else []
    rh_vib_zones = RH_VIB_ZONES if is_rh_np and test_key == "step4" else []
    rh_second_impact_zones = RH_SECOND_IMPACT_ZONES if is_rh_np and test_key == "step5" else []
    rh_step12_zones = RH_STEP12_ZONES if is_rh_np and test_key == "step12" else []

    # --- Xử lý upload ảnh, xóa ảnh, comment, status ---
    if request.method == 'POST':
        # Chỉ upload ảnh loại thường (test_imgs)
        if 'test_imgs' in request.files:
            files = request.files.getlist('test_imgs')
            for file in files:
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[-1].lower()
                    prefix = f"test_{group}_{test_key}_"
                    current_nums = [
                        int(f[len(prefix):].split('.')[0])
                        for f in os.listdir(report_folder)
                        if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()
                    ]
                    next_num = max(current_nums) + 1 if current_nums else 1
                    new_fname = f"{prefix}{next_num}.{ext}"
                    file.save(os.path.join(report_folder, new_fname))
        # Xóa ảnh thường
        if 'delete_img' in request.form:
            del_img = request.form['delete_img']
            img_path = os.path.join(report_folder, del_img)
            if os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass
        # Ghi status PASS/FAIL/N/A
        if 'status' in request.form:
            update_group_note_file(status_file, test_key, request.form['status'])  # DÙNG SAFE
        # Ghi comment
        if 'save_comment' in request.form:
            comment_val = request.form.get('comment_input', '').strip()
            update_group_note_file(comment_file, test_key, comment_val)  # DÙNG SAFE
        return redirect(request.url)

    # --- Chuẩn bị dữ liệu ảnh vùng RH (step3/4/5/12) ---
    zone_imgs = {}
    for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones + rh_step12_zones:
        imgs_zone = []
        for f in os.listdir(report_folder):
            if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_{zone}_"):
                imgs_zone.append(f"/images/{report}/{f}")
        zone_imgs[zone] = imgs_zone

    # --- Chuẩn bị dữ liệu ảnh thường ---
    imgs = []
    for f in sorted(os.listdir(report_folder)):
        # Chỉ lấy ảnh loại thường, không lấy ảnh vùng
        if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_") and all(not f.startswith(f"test_{group}_{test_key}_{zone}_") for zone, _ in rh_impact_zones + rh_vib_zones + rh_second_impact_zones + rh_step12_zones):
            imgs.append(f"/images/{report}/{f}")

    # --- Chuẩn bị ảnh drop, impact, rot nếu có ---
    drop_imgs, impact_imgs, rot_imgs = [], [], []
    if is_drop:
        for zone in DROP_ZONES:
            di = []
            for f in os.listdir(report_folder):
                if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_drop_{zone}_"):
                    di.append(f"/images/{report}/{f}")
            drop_imgs.append(di)
    if is_impact:
        for zone in IMPACT_ZONES:
            ii = []
            for f in os.listdir(report_folder):
                if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_impact_{zone}_"):
                    ii.append(f"/images/{report}/{f}")
            impact_imgs.append(ii)
    if is_rot:
        for zone in ROT_ZONES:
            ri = []
            for f in os.listdir(report_folder):
                if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_rotation_{zone}_"):
                    ri.append(f"/images/{report}/{f}")
            rot_imgs.append(ri)

    # --- Trả về template ---
    return render_test_group_item(report, group, test_key, group_titles, comment=comment)

def render_test_group_item(report, group, key, group_titles, comment):
    title = group_titles[key]
    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(report_folder, exist_ok=True)
    status_file = os.path.join(report_folder, f"status_{group}.txt")
    comment_file = os.path.join(report_folder, f"comment_{group}.txt")

    # RH Non Pallet zone logic
    is_rh_np = (group == "transit_RH_np")
    is_rh_np_step3 = is_rh_np and key == "step3"
    is_rh_np_step4 = is_rh_np and key == "step4"
    is_rh_np_step5 = is_rh_np and key == "step5"
    rh_impact_zones = RH_IMPACT_ZONES if is_rh_np_step3 else []
    rh_vib_zones = RH_VIB_ZONES if is_rh_np_step4 else []
    rh_second_impact_zones = RH_SECOND_IMPACT_ZONES if is_rh_np_step5 else []
    allow_na = is_rh_np and (key in ["step6", "step7", "step8", "step11", "step12"])

    # Xử lý ảnh vùng RH (zone_imgs)
    zone_imgs = {}
    for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
        imgs = []
        if os.path.exists(report_folder):
            for f in os.listdir(report_folder):
                if allowed_file(f) and f.startswith(f"test_{group}_{key}_{zone}_"):
                    imgs.append(f"/images/{report}/{f}")
        zone_imgs[zone] = imgs

    # Vùng Face cho transit_181_gt68 step4
    gt68_face_zones, gt68_face_labels, gt68_face_imgs = [], [], []
    if group == "transit_181_gt68" and key == "step4":
        gt68_face_zones = GT68_FACE_ZONES
        gt68_face_labels = GT68_FACE_LABELS
        for zone in gt68_face_zones:
            imgs = []
            if os.path.exists(report_folder):
                for f in os.listdir(report_folder):
                    if allowed_file(f) and f.startswith(f"test_{group}_{key}_gt68_face_{zone}_"):
                        imgs.append(f"/images/{report}/{f}")
            gt68_face_imgs.append(imgs)

    # Nhóm transit 2C logic
    TRANSIT_2C_GROUPS = ("transit_2c_np", "transit_2c_pallet")
    if not (is_rh_np_step3 or is_rh_np_step4 or is_rh_np_step5):
        is_transit_2c = group in TRANSIT_2C_GROUPS
        is_drop = (is_drop_test(title) and is_transit_2c) or (group == "transit_181_lt68" and key == "step4")
        is_impact = is_impact_test(title) and is_transit_2c
        is_rot = is_rotational_test(title) and is_transit_2c
    else:
        is_drop = is_impact = is_rot = False

    # Drop, Impact, Rot imgs
    drop_imgs = []
    if is_drop:
        for zone in DROP_ZONES:
            imgs = []
            if os.path.exists(report_folder):
                for f in os.listdir(report_folder):
                    if allowed_file(f) and f.startswith(f"test_{group}_{key}_drop_{zone}_"):
                        imgs.append(f"/images/{report}/{f}")
            drop_imgs.append(imgs)

    impact_imgs = []
    rot_imgs = []
    if is_impact:
        for zone in IMPACT_ZONES:
            imgs = []
            if os.path.exists(report_folder):
                for f in os.listdir(report_folder):
                    if allowed_file(f) and f.startswith(f"test_{group}_{key}_impact_{zone}_"):
                        imgs.append(f"/images/{report}/{f}")
            impact_imgs.append(imgs)
    if is_rot:
        for zone in ROT_ZONES:
            imgs = []
            if os.path.exists(report_folder):
                for f in os.listdir(report_folder):
                    if allowed_file(f) and f.startswith(f"test_{group}_{key}_rotation_{zone}_"):
                        imgs.append(f"/images/{report}/{f}")
            rot_imgs.append(imgs)
    gt68_face_imgs = []
    if group == "transit_181_gt68" and key == "step4":
        for zone in GT68_FACE_ZONES:
            imgs = []
            if os.path.exists(report_folder):
                for f in os.listdir(report_folder):
                    if allowed_file(f) and f.startswith(f"test_{group}_{key}_gt68_face_{zone}_"):
                        imgs.append(f"/images/{report}/{f}")
            gt68_face_imgs.append(imgs)

    # === Status/comment helper ===
    def update_group_note_file(file_path, key, value):
        # Đọc file dùng lock
        lines = []
        found = False
        content = safe_read_text(file_path)
        if content:
            lines = content.splitlines(keepends=True)
        new_lines = []
        for line in lines:
            if line.strip().startswith(f"Mục {key}:"):
                new_lines.append(f"Mục {key}: {value}\n")
                found = True
            else:
                new_lines.append(line)
        if not found:
            new_lines.append(f"Mục {key}: {value}\n")
        # Ghi lại dùng lock
        safe_write_text(file_path, "".join(new_lines))

    def get_group_note_value(file_path, key):
        content = safe_read_text(file_path)
        if content:
            for line in content.splitlines():
                s = line.strip()
                if s.startswith(f"Mục {key}:") or s.startswith(f"{key}:"):
                    return s.split(":", 1)[1].strip()
        return None

    status_value = get_group_note_value(status_file, key)

    # === Xử lý POST: chỉ xử lý xóa ảnh, status, comment (KHÔNG UPLOAD ẢNH VÙNG ZONE Ở ĐÂY) ===
    if request.method == 'POST':
        # Xóa ảnh thường hoặc vùng
        if 'delete_img' in request.form:
            del_img = request.form['delete_img']
            img_path = os.path.join(report_folder, del_img)
            if os.path.exists(img_path):
                os.remove(img_path)
        # Ghi status PASS/FAIL/N/A
        if 'status' in request.form:
            status = request.form['status']
            update_group_note_file(status_file, key, status)
        # Ghi comment
        if 'save_comment' in request.form:
            comment = request.form.get('comment_input', '').strip()
            update_group_note_file(comment_file, key, comment)
        # Cập nhật loại kiểm tra gần nhất
        vi_name = TEST_TYPE_VI.get(group, group.upper())
        set_last_test_type(report, vi_name)
        return redirect(request.url)

    # === Tính trạng thái tổng thể từng mục cho menu group ===
    test_status = {}
    for k in group_titles:
        st = get_group_note_value(status_file, k) if not group.startswith("transit") else None
        cm = get_group_note_value(comment_file, k)
        has_img = False
        prefix = f"test_{group}_{k}_"
        if os.path.exists(report_folder):
            for fn in os.listdir(report_folder):
                if allowed_file(fn) and fn.startswith(prefix):
                    has_img = True
                    break
        test_status[k] = {
            'status': st,
            'comment': cm,
            'has_img': has_img
        }

    # === Lấy ảnh thường cho mục không phải drop/impact/rot/RH np ===
    imgs = []
    if os.path.exists(report_folder) and not is_drop:
        prefix = f"test_{group}_{key}_"
        for f in sorted(os.listdir(report_folder)):
            if allowed_file(f) and f.startswith(prefix):
                # Chỉ nhận file có số thứ tự ngay sau prefix (vd: ..._1.jpg, ..._2.png)
                tail = f[len(prefix):].split('.')[0]
                if tail.isdigit():
                    imgs.append(f"/images/{report}/{f}")

    # === Chọn template (transit dùng test_transit_item.html) ===
    TRANSIT_GROUPS = (
        "transit_2c_np", "transit_2c_pallet",
        "transit_RH_np", "transit_RH_pallet",
        "transit_181_lt68", "transit_181_gt68",
        "transit_3b_np", "transit_3b_pallet", "transit_3a_np"
    )
    if group in TRANSIT_GROUPS:
        template_name = "test_transit_item.html"
    else:
        template_name = "test_group_item.html"

    return render_template(
        template_name,
        report=report,
        group=group,
        test_titles=group_titles,
        test_status=test_status,
        key=key,
        imgs=imgs,
        status=status_value,
        comment=get_group_note_value(comment_file, key),
        title=title,
        is_drop=is_drop,
        drop_labels=DROP_LABELS,
        drop_zones=DROP_ZONES,
        drop_imgs=drop_imgs,
        is_impact=is_impact,
        impact_labels=IMPACT_LABELS,
        impact_zones=IMPACT_ZONES,
        impact_imgs=impact_imgs,
        is_rot=is_rot,
        rot_labels=ROT_LABELS,
        rot_zones=ROT_ZONES,
        rot_imgs=rot_imgs,
        is_rh_np=is_rh_np,
        rh_impact_zones=rh_impact_zones,
        rh_vib_zones=rh_vib_zones,
        rh_second_impact_zones=rh_second_impact_zones,
        allow_na=allow_na,
        zone_imgs=zone_imgs,
        gt68_face_labels=GT68_FACE_LABELS,
        gt68_face_zones=GT68_FACE_ZONES,
        gt68_face_imgs=gt68_face_imgs,
    )

def read_kv_file(path):
    """
    Đọc file dạng:
        key: value
        key2: value2
    -> trả về dict {key: value}
    """
    data = {}
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                for raw in f.readlines():
                    line = raw.strip()
                    if not line:
                        continue
                    # chấp nhận cả ':' lẫn '=' cho bền
                    if ":" in line:
                        k, v = line.split(":", 1)
                    elif "=" in line:
                        k, v = line.split("=", 1)
                    else:
                        # nếu là file kiểu cũ chỉ có 1 giá trị (PASS/FAIL/...)
                        # gán vào key 'default'
                        data["default"] = line
                        continue
                    data[k.strip()] = v.strip()
    except Exception:
        pass
    return data


def upsert_kv_line(path, key, value):
    """
    Ghi/ cập nhật một dòng 'key: value' vào file.
    - Nếu chưa có file => tạo mới.
    - Nếu có => cập nhật đúng key, giữ các key khác.
    """
    d = read_kv_file(path)
    d[key] = value
    try:
        with open(path, "w", encoding="utf-8") as f:
            for k, v in d.items():
                f.write(f"{k}: {v}\n")
    except Exception:
        pass


# ====== Route hot_cold (ghi status/comment theo từng test_key) ===============
# Cho phép URL có/không có test_key (mặc định là 'hot_cold' để không phá link cũ)
@app.route("/hot_cold_test/<report>/<group>", defaults={'test_key': 'hot_cold'}, methods=["GET", "POST"])
@app.route("/hot_cold_test/<report>/<group>/<test_key>", methods=["GET", "POST"])
def hot_cold_test(report, group, test_key):
    from_line = request.args.get("from_line")

    # ====== Lấy tên hiển thị đúng theo test_key ======
    try:
        raw_title = TEST_GROUP_TITLES.get(group, {}).get(test_key)
    except Exception:
        raw_title = None

    if isinstance(raw_title, dict):
        display_title = raw_title.get('short') or raw_title.get('full') or test_key.replace('_', ' ').title()
    elif raw_title:
        display_title = raw_title
    else:
        display_title = test_key.replace('_', ' ').title()

    set_last_test_type(report, f"{display_title} ({group.upper()})")

    # ====== Chuẩn bị đường dẫn/lưu trữ ======
    vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(folder, exist_ok=True)

    # File CHUNG theo group (mỗi nút 1 dòng)
    common_prefix = f"{group}"
    status_file   = os.path.join(folder, f"status_{common_prefix}.txt")
    comment_file  = os.path.join(folder, f"comment_{common_prefix}.txt")

    # Ảnh & mốc thời gian vẫn theo test_key (không đổi)
    test_prefix       = f"{test_key}_{group}"   # ví dụ: hot_cold_indoor_thuong
    before_tag        = f"{test_key}_before_{group}"
    after_tag         = f"{test_key}_after_{group}"
    before_time_file  = os.path.join(folder, f"{test_prefix}_before_time.txt")
    after_time_file   = os.path.join(folder, f"{test_prefix}_after_time.txt")
    duration_file     = os.path.join(folder, f"{test_prefix}_duration.txt")

    # ====== Xử lý POST ======
    if request.method == "POST":
        # 1) Cập nhật trạng thái -> ghi/upsert theo test_key (KHÔNG ghi đè cả file)
        if "status" in request.form:
            status_value = (request.form.get("status") or "").strip()
            if status_value:
                upsert_kv_line(status_file, test_key, status_value)

        # 2) Lưu ghi chú -> cũng upsert theo test_key
        if "save_comment" in request.form:
            cmt = (request.form.get("comment_input") or "").strip()
            # để tránh phá format một dòng, thay newline bằng ' / '
            cmt_one_line = " / ".join([s.strip() for s in cmt.splitlines() if s.strip()]) if cmt else ""
            upsert_kv_line(comment_file, test_key, cmt_one_line)

        # 3) Upload ảnh (before/after) + ghi mốc thời gian tương ứng
        for tag, time_file in [(before_tag, before_time_file), (after_tag, after_time_file)]:
            field_name = f"{tag}_imgs"
            if field_name in request.files:
                files = request.files.getlist(field_name)
                count = 0
                for file in files:
                    if file and allowed_file(file.filename):
                        ext = file.filename.rsplit('.', 1)[-1].lower()
                        prefix_img = f"{tag}_"
                        nums = [
                            int(fname[len(prefix_img):].split('.')[0])
                            for fname in os.listdir(folder)
                            if fname.startswith(prefix_img) and fname[len(prefix_img):].split('.')[0].isdigit()
                        ]
                        next_num = max(nums) + 1 if nums else 1
                        new_fname = f"{tag}_{next_num}.{ext}"
                        file.save(os.path.join(folder, new_fname))
                        count += 1
                if count > 0:
                    now = datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M")
                    safe_write_text(time_file, now)

        # 4) Xoá ảnh (và dọn mốc thời gian nếu hết ảnh)
        if "delete_img" in request.form:
            img = request.form["delete_img"]
            img_path = os.path.join(folder, img)
            if os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except Exception:
                    pass
            # Nếu không còn ảnh before/after thì xoá file time tương ứng
            if img.startswith(before_tag):
                still = [f for f in os.listdir(folder) if allowed_file(f) and f.startswith(before_tag)]
                if not still and os.path.exists(before_time_file):
                    try: os.remove(before_time_file)
                    except Exception: pass
            if img.startswith(after_tag):
                still = [f for f in os.listdir(folder) if allowed_file(f) and f.startswith(after_tag)]
                if not still and os.path.exists(after_time_file):
                    try: os.remove(after_time_file)
                    except Exception: pass

        # 5) Cập nhật thời gian test (giờ)
        if "set_duration" in request.form:
            raw = (request.form.get("duration") or "").strip()
            try:
                dur = float(raw)
                if dur <= 0: raise ValueError
                safe_write_text(duration_file, str(dur))
                flash("Đã cập nhật thời gian test.", "success")
            except Exception:
                flash("Giá trị thời gian không hợp lệ.", "danger")

        # tránh resubmit
        session[f"last_test_type_{report}"] = f"{display_title} ({group.upper()})"
        return redirect(request.url)

    # ====== Đọc dữ liệu để render (lấy đúng mục theo test_key) ======
    status_map  = read_kv_file(status_file)
    comment_map = read_kv_file(comment_file)

    # Ưu tiên key cụ thể; nếu trước đây file cũ lưu 1 dòng không key thì dùng 'default'
    status  = (status_map.get(test_key) or status_map.get("default") or "").strip()
    comment = (comment_map.get(test_key) or comment_map.get("default") or "").strip()

    # Hình mô tả (nếu có trong TEST_GROUP_TITLES)
    try:
        imgs_mo_ta = (TEST_GROUP_TITLES.get(group, {}).get(test_key) or {}).get("img", [])
    except Exception:
        imgs_mo_ta = []

    # Danh sách ảnh before/after
    imgs_before, imgs_after = [], []
    for fname in sorted(os.listdir(folder)):
        if allowed_file(fname):
            if fname.startswith(before_tag):
                imgs_before.append(f"/images/{report}/{fname}")
            elif fname.startswith(after_tag):
                imgs_after.append(f"/images/{report}/{fname}")

    # Thời gian upload
    before_upload_time = (safe_read_text(before_time_file) or "").strip() if os.path.exists(before_time_file) else None
    after_upload_time  = (safe_read_text(after_time_file) or "").strip()  if os.path.exists(after_time_file)  else None

    # Thời lượng test (giờ)
    raw_duration = safe_read_text(duration_file)
    try:
        so_gio_test = float(raw_duration) if raw_duration not in (None, "") else float(SO_GIO_TEST)
    except Exception:
        so_gio_test = 4.0

    # ====== Render ======
    return render_template(
        "hot_cold_test.html",
        report=report,
        group=group,
        test_key=test_key,
        title={'short': display_title, 'full': display_title},
        status=status,
        comment=comment,
        imgs_mo_ta=imgs_mo_ta,
        imgs_before=imgs_before,
        imgs_after=imgs_after,
        before_upload_time=before_upload_time,
        after_upload_time=after_upload_time,
        so_gio_test=so_gio_test,
        from_line=from_line,
        before_tag=before_tag,
        after_tag=after_tag,
    )

def get_hotcold_elapsed(report, group):
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    time_file = os.path.join(folder, f"hot_cold_upload_time_{group}.txt")
    tstr = safe_read_text(time_file).strip() if os.path.exists(time_file) else ""
    if tstr:
        try:
            vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
            dt = datetime.strptime(tstr, "%d/%m/%Y %H:%M")
            dt = vn_tz.localize(dt)
            now = datetime.now(vn_tz)
            return (now - dt).total_seconds() / 3600
        except Exception:
            return None
    return None

@app.route("/line_test/<report>", methods=["GET", "POST"])
def line_test(report):
    set_last_test_type(report, "LINE TEST")
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(folder, exist_ok=True)
    before_tag, after_tag = "line_before", "line_after"
    files_map = {
        "status": os.path.join(folder, "line_status.txt"),
        "comment": os.path.join(folder, "line_comment.txt"),
        "before_time": os.path.join(folder, "before_upload_time.txt"),
        "after_time": os.path.join(folder, "after_upload_time.txt"),
    }
    fail_reasons_list = [
        "Vật liệu bị ẩm.",
        "Vị trí bị tách lớp, mặt dưới veneer có phủ keo.",
        "Vị trí bị tách lớp, mặt dưới veneer không phủ đều keo."
    ]

    # --- POST ---
    if request.method == "POST":
        # Lưu trạng thái PASS/FAIL/DATA
        if "status" in request.form:
            safe_write_text(files_map["status"], request.form["status"])
            if request.form["status"] != "FAIL":
                if os.path.exists(files_map["comment"]):
                    os.remove(files_map["comment"])
        # Lưu fail reason
        if "save_fail_reason" in request.form:
            reasons = request.form.getlist("fail_reason")
            other = request.form.get("fail_reason_other", "").strip()
            if other: reasons.append(other)
            safe_write_text(files_map["comment"], "; ".join(reasons))
        # Upload ảnh before/after
        for tag, time_file in [(before_tag, files_map["before_time"]), (after_tag, files_map["after_time"])]:
            if f"{tag}_imgs" in request.files:
                files = request.files.getlist(f"{tag}_imgs")
                nums = [int(f[len(tag)+1:].split('.')[0]) for f in os.listdir(folder)
                        if f.startswith(f"{tag}_") and f[len(tag)+1:].split('.')[0].isdigit()]
                next_num = max(nums, default=0) + 1
                count = 0
                for file in files:
                    if file and allowed_file(file.filename):
                        ext = file.filename.rsplit('.', 1)[-1].lower()
                        file.save(os.path.join(folder, f"{tag}_{next_num}.{ext}"))
                        next_num += 1
                        count += 1
                if count:
                    vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
                    safe_write_text(time_file, datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M"))
        # Xóa ảnh
        if "delete_img" in request.form:
            img = request.form["delete_img"]
            img_path = os.path.join(folder, img)
            if os.path.exists(img_path): os.remove(img_path)
            for tag, time_file in [(before_tag, files_map["before_time"]), (after_tag, files_map["after_time"])]:
                if img.startswith(tag):
                    if not any(allowed_file(f) and f.startswith(tag) for f in os.listdir(folder)):
                        if os.path.exists(time_file): os.remove(time_file)
        set_last_test_type(report, "LINE TEST")
        return redirect(request.url)

    # --- GET: Đọc dữ liệu đã lưu ---
    status = safe_read_text(files_map["status"])
    fail_reason_raw = safe_read_text(files_map["comment"])
    fail_reasons, fail_reason_other = [], ""
    if fail_reason_raw:
        all_reasons = [r.strip() for r in fail_reason_raw.split(";") if r.strip()]
        for r in all_reasons[:]:
            if r not in fail_reasons_list:
                fail_reason_other = r
                all_reasons.remove(r)
        fail_reasons = all_reasons
    imgs_before = [f"/images/{report}/{f}" for f in sorted(os.listdir(folder)) if allowed_file(f) and f.startswith(before_tag)]
    imgs_after  = [f"/images/{report}/{f}" for f in sorted(os.listdir(folder)) if allowed_file(f) and f.startswith(after_tag)]
    before_upload_time = safe_read_text(files_map["before_time"])
    after_upload_time  = safe_read_text(files_map["after_time"])

    return render_template(
        "line_test.html",
        report=report,
        status=status,
        fail_reasons=fail_reasons,
        fail_reason_other=fail_reason_other,
        fail_reasons_list=fail_reasons_list,
        imgs_before=imgs_before,
        imgs_after=imgs_after,
        before_upload_time=before_upload_time,
        after_upload_time=after_upload_time,
        so_gio_test=SO_GIO_TEST,
    )

def get_line_test_elapsed(report):
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    before_time_file = os.path.join(folder, "before_upload_time.txt")
    tstr = safe_read_text(before_time_file)  # Dùng hàm an toàn, đã có filelock
    if tstr:
        try:
            vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
            dt = datetime.strptime(tstr, "%d/%m/%Y %H:%M")
            dt = vn_tz.localize(dt)
            now = datetime.now(vn_tz)
            elapsed = (now - dt).total_seconds() / 3600
            return elapsed
        except Exception as e:
            print("Parse time error:", e)
            return None
    return None

SAMPLE_STORAGE_FILE = os.path.join(BASE_DIR, "sample_storage.json")

@app.route("/store_sample", methods=["GET", "POST"])
def store_sample():
    """
    Route lưu mẫu theo format mới (box_code, save_date, discard_date...),
    cho phép nhiều mẫu trong cùng 1 box (list dạng).
    """
    report = request.args.get("report")
    item_code = get_item_code(report)
    error_msg = ""

    # --- Đọc dữ liệu an toàn ---
    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    if not isinstance(SAMPLE_STORAGE, dict):
        SAMPLE_STORAGE = {}

    # --- Kiểm tra xem mẫu đã tồn tại chưa (theo report + item_code) ---
    existing_box = None
    for box, info in SAMPLE_STORAGE.items():
        if isinstance(info, list):
            # box chứa nhiều mẫu
            for entry in info:
                if entry.get("report") == report and entry.get("item_code") == item_code:
                    existing_box = box
                    break
        elif isinstance(info, dict):
            # dữ liệu cũ chỉ có 1 mẫu
            if info.get("report") == report and info.get("item_code") == item_code:
                existing_box = box
        if existing_box:
            break

    # Nếu đã tồn tại mẫu => chuyển sang trang info
    if existing_box:
        return redirect(url_for("sample_infor", report=report, box_code=existing_box))

    # --- Nếu chưa có mẫu => xử lý form POST ---
    if request.method == "POST":
        sample_type = request.form.get("sample_type", "").strip()
        box_code = request.form.get("box_code", "").strip().upper()
        note = request.form.get("note", "").strip()

        save_date_str = request.form.get("save_date")
        discard_date_str = request.form.get("discard_date")

        from datetime import datetime, date
        from dateutil.relativedelta import relativedelta

        try:
            save_date = datetime.strptime(save_date_str, "%Y-%m-%d").date() if save_date_str else date.today()
        except:
            save_date = date.today()

        try:
            discard_date = datetime.strptime(discard_date_str, "%Y-%m-%d").date() if discard_date_str else (save_date + relativedelta(months=3))
        except:
            discard_date = save_date + relativedelta(months=3)

        # --- Ghi dữ liệu ---
        if box_code not in SAMPLE_STORAGE:
            SAMPLE_STORAGE[box_code] = []
        elif isinstance(SAMPLE_STORAGE[box_code], dict):
            # nếu dữ liệu cũ chỉ có 1 mẫu thì chuyển thành list
            SAMPLE_STORAGE[box_code] = [SAMPLE_STORAGE[box_code]]

        # thêm mẫu mới vào danh sách trong box
        SAMPLE_STORAGE[box_code].append({
            "report": report,
            "item_code": item_code,
            "sample_type": sample_type,
            "box_code": box_code,
            "save_date": save_date.isoformat(),
            "discard_date": discard_date.isoformat(),
            "note": note
        })

        safe_write_json(SAMPLE_STORAGE_FILE, SAMPLE_STORAGE)
        return redirect(url_for("sample_infor", report=report, box_code=box_code))

    # --- GET request: render form ---
    return render_template(
        "sample_form.html",
        report=report,
        item_code=item_code
    )

@app.route("/report/<report>/sample/edit/<box_code>", methods=["GET", "POST"])
def edit_sample(report, box_code):
    """Chỉnh sửa mẫu có sẵn trong box (tương thích nhiều mẫu)."""
    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    if not isinstance(SAMPLE_STORAGE, dict):
        SAMPLE_STORAGE = {}

    data = None
    if box_code in SAMPLE_STORAGE:
        box_data = SAMPLE_STORAGE[box_code]
        if isinstance(box_data, list):
            for entry in box_data:
                if entry.get("report") == report:
                    data = entry
                    break
        elif isinstance(box_data, dict):
            data = box_data

    if not data:
        return f"Không tìm thấy mẫu trong box {box_code}", 404

    from datetime import datetime, date
    from dateutil.relativedelta import relativedelta

    if request.method == "POST":
        sample_type = request.form.get("sample_type")
        note = request.form.get("note", "")
        save_date_str = request.form.get("save_date")
        discard_date_str = request.form.get("discard_date")

        try:
            save_date = datetime.strptime(save_date_str, "%Y-%m-%d").date()
        except:
            save_date = date.today()

        try:
            discard_date = datetime.strptime(discard_date_str, "%Y-%m-%d").date()
        except:
            discard_date = save_date + relativedelta(months=3)

        # cập nhật dữ liệu trong list đúng mẫu đó
        if isinstance(SAMPLE_STORAGE[box_code], list):
            for i, entry in enumerate(SAMPLE_STORAGE[box_code]):
                if entry.get("report") == report:
                    SAMPLE_STORAGE[box_code][i].update({
                        "sample_type": sample_type,
                        "note": note,
                        "save_date": save_date.isoformat(),
                        "discard_date": discard_date.isoformat()
                    })
                    break
        elif isinstance(SAMPLE_STORAGE[box_code], dict):
            SAMPLE_STORAGE[box_code].update({
                "sample_type": sample_type,
                "note": note,
                "save_date": save_date.isoformat(),
                "discard_date": discard_date.isoformat()
            })

        safe_write_json(SAMPLE_STORAGE_FILE, SAMPLE_STORAGE)
        return redirect(url_for("list_samples", report=report))

    # --- GET request: render form ---
    return render_template(
        "sample_form.html",
        report=report,
        item_code=data.get("item_code", ""),
        box_code=box_code,
        sample_type=data.get("sample_type", ""),
        note=data.get("note", ""),
        save_date=data.get("save_date", ""),
        discard_date=data.get("discard_date", "")
    )

@app.route("/boxes", methods=["GET","POST","DELETE"])
def boxes_api():
    boxes = read_boxes()

    if request.method == "GET":
        return jsonify({"boxes": boxes})

    if request.method == "POST":
        payload = request.get_json(silent=True) or {}
        code = (payload.get("code") or "").strip().upper()
        if code and code not in boxes:
            boxes.append(code)
            boxes = sorted(set(boxes))
            write_boxes(boxes)
        return jsonify({"boxes": boxes})

    if request.method == "DELETE":
        code = (request.args.get("code") or "").strip().upper()
        if code in boxes:
            boxes = [b for b in boxes if b != code]
            write_boxes(boxes)
        return jsonify({"boxes": boxes})
    
@app.route("/report/<report>/sample/info/<box_code>")
def sample_infor(report, box_code):
    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    info = None
    if box_code in SAMPLE_STORAGE:
        box_data = SAMPLE_STORAGE[box_code]
        if isinstance(box_data, list):
            # tìm mẫu đúng report trong box
            for entry in box_data:
                if entry.get("report") == report:
                    info = entry
                    break
        elif isinstance(box_data, dict):
            info = box_data
    if not info:
        return f"Không tìm thấy mẫu trong box {box_code}", 404

    return render_template(
        "sample_infor.html",
        info=info,
        report_id=report,
        item_code=info.get("item_code", ""),
        box_code=box_code
    )

@app.route("/list_samples", methods=["GET"])
def list_samples():
    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    if not isinstance(SAMPLE_STORAGE, dict):
        SAMPLE_STORAGE = {}

    samples = []

    for box_code, info in SAMPLE_STORAGE.items():
        if isinstance(info, list):
            # box có nhiều mẫu
            for entry in info:
                samples.append({
                    "box_code": box_code,
                    "report": entry.get("report", ""),
                    "item_code": entry.get("item_code", ""),
                    "sample_type": entry.get("sample_type", ""),
                    "save_date": entry.get("save_date", ""),
                    "discard_date": entry.get("discard_date", ""),
                    "note": entry.get("note", "")
                })
        elif isinstance(info, dict):
            # dữ liệu cũ chỉ có 1 mẫu
            samples.append({
                "box_code": box_code,
                "report": info.get("report", ""),
                "item_code": info.get("item_code", ""),
                "sample_type": info.get("sample_type", ""),
                "save_date": info.get("save_date", ""),
                "discard_date": info.get("discard_date", ""),
                "note": info.get("note", "")
            })

    # Debug log (tùy chọn)
    print(f"✅ Loaded {len(samples)} samples from {len(SAMPLE_STORAGE)} boxes")

    # Truyền JSON vào template để render
    from markupsafe import Markup
    import json
    samples_json = Markup(json.dumps(samples, ensure_ascii=False))

    return render_template(
        "list_samples.html",
        samples_json=samples_json,
        report_id=request.args.get("report", "")
    )

# ====== API: XÓA MẪU THEO BOX_CODE ======
@app.route("/samples", methods=["DELETE"])
def delete_sample():
    box_code = (request.args.get("box_code") or "").strip().upper()
    report = (request.args.get("report") or "").strip()

    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    if not isinstance(SAMPLE_STORAGE, dict):
        SAMPLE_STORAGE = {}

    if box_code not in SAMPLE_STORAGE:
        return {"ok": False, "error": "not found"}, 404

    box_data = SAMPLE_STORAGE[box_code]

    if isinstance(box_data, list) and report:
        box_data = [r for r in box_data if r.get("report") != report]
        if box_data:
            SAMPLE_STORAGE[box_code] = box_data
        else:
            del SAMPLE_STORAGE[box_code]
    else:
        del SAMPLE_STORAGE[box_code]

    safe_write_json(SAMPLE_STORAGE_FILE, SAMPLE_STORAGE)
    return {"ok": True}

@app.post('/export_expired_samples')
def export_expired_samples():
    """Nhận list mẫu từ client, lọc các mẫu hết hạn (đang có status-danger) để xuất Excel."""
    data = request.get_json(force=True) or []
    # lọc những mẫu có ngày hủy < hôm nay
    today = datetime.now().date()
    expired = [r for r in data if r.get('discard_date') and r.get('discard_date') < str(today)]

    wb_bytes = export_expired_samples_to_excel(expired)
    return send_file(
        BytesIO(wb_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='expired_samples.xlsx'
    )

@app.route('/images/<report>/imgs_<group>_<test_key>/<filename>')
def serve_test_img(report, group, test_key, filename):
    folder = os.path.join(UPLOAD_FOLDER, report, f"imgs_{group}_{test_key}")
    if not os.path.exists(folder):
        # Báo lỗi rõ ràng hoặc trả về 404
        return "Không tìm thấy thư mục ảnh!", 404
    return send_from_directory(folder, filename)

@app.route("/api/report_comment")
def api_report_comment():
    report = (request.args.get("report") or "").strip()
    if not report:
        return jsonify({"comment": ""})

    base_dir = os.path.dirname(os.path.abspath(__file__))
    rep_dir = os.path.join(base_dir, "images", report)
    if not os.path.isdir(rep_dir):
        return jsonify({"comment": ""})

    # Ưu tiên file có chữ "comment" và đuôi .txt/.md/.json/.log; fallback các .txt khác
    patterns = ["comment*.txt", "comment*.md", "comment*.json", "comment*.log",
                "*.txt", "*.md", "*.json"]
    candidates = []
    for pat in patterns:
        candidates.extend(glob.glob(os.path.join(rep_dir, pat)))

    # Sắp theo ưu tiên 'comment' rồi mới đến thời gian mới nhất
    candidates.sort(key=lambda p: (0 if "comment" in os.path.basename(p).lower() else 1, -os.path.getmtime(p)))

    text = ""
    for p in candidates:
        try:
            if p.lower().endswith(".json"):
                import json
                with open(p, "r", encoding="utf-8", errors="ignore") as f:
                    obj = json.load(f)
                # chấp nhận khóa 'comment' hoặc toàn bộ nội dung là chuỗi
                if isinstance(obj, dict) and "comment" in obj:
                    text = str(obj["comment"])
                elif isinstance(obj, str):
                    text = obj
                else:
                    # gộp các value thành text
                    text = "\n".join(f"{k}: {v}" for k, v in obj.items())
            else:
                with open(p, "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()
            if text.strip():
                break
        except Exception:
            continue

    return jsonify({"comment": text})

@app.route("/view_counter_log")
def view_counter_log():

    excel_path = "counter_detail_log.xlsx"
    rows = []
    type_of_set = set()
    ca_map = {"office": "HC", "hc": "HC", "ot": "OT"}
    header = ["Ngày", "Ca", "Tổng"]  # Default

    if os.path.exists(excel_path):
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
            ws = wb.active
            # Build column name -> index map
            col_idx = {str(cell.value).strip().lower(): i for i, cell in enumerate(ws[1], 0)}
            date_idx = col_idx.get("ngày", 0)
            ca_idx = col_idx.get("ca", 2)
            type_idx = col_idx.get("type of", 4)

            # summary[day][ca][type_of_short] = count
            summary = OrderedDict()
            for row in ws.iter_rows(min_row=2, values_only=True):
                day = row[date_idx]
                ca_raw = str(row[ca_idx]).strip().lower() if row[ca_idx] else ""
                ca = "HC" if "office" in ca_raw or ca_raw == "hc" else "OT"
                type_of_raw = (row[type_idx] or "UNKNOWN").strip().upper()
                if type_of_raw.startswith("OUTSOURCE"):
                    parts = type_of_raw.split("-")
                    type_of_short = (parts[-1][-3:] if len(parts) >= 2 else type_of_raw[-3:])
                else:
                    type_of_short = type_of_raw[:3]
                type_of_short = type_of_raw[:3]
                type_of_set.add(type_of_short)
                if day not in summary:
                    summary[day] = {"HC": defaultdict(int), "OT": defaultdict(int)}
                summary[day][ca][type_of_short] += 1

            # Giữ 10 ngày mới nhất
            day_keys = list(summary.keys())[-10:]
            summary = OrderedDict((k, summary[k]) for k in day_keys)
            type_of_list = sorted([t for t in type_of_set if t != "UNK"])
            if "UNK" in type_of_set:
                type_of_list.append("UNK")
            header = ["Ngày", "Ca"] + type_of_list + ["Tổng"]

            # Tạo rows cho template (2 dòng/ngày: HC, OT)
            rows = []
            for day in summary:
                for ca in ("HC", "OT"):
                    type_counts = [summary[day][ca].get(t, 0) for t in type_of_list]
                    rows.append({
                        "date": day if ca == "HC" else "",
                        "ca": ca,
                        "types": type_counts,
                        "total": sum(type_counts)
                    })
        except Exception as e:
            # Log lỗi nếu cần, nhưng trả template bình thường
            print("[view_counter_log] Error:", e)
            rows = []
            type_of_list = []
    else:
        type_of_list = []

    return render_template(
        "counter_log.html",
        header=header,
        rows=rows,
        type_of_list=type_of_list,
    )

DISPLAY = {
    "hot_cold": "Hot & Cold cycle test",
    "standing_water": "Standing water test",
    "stain": "Stain test",
    "corrosion": "Corrosion test",
}

@app.route("/set_pref", methods=["POST"])
def set_pref():
    key = (request.json.get("key") or "").strip()
    value = str(request.json.get("value") or "").strip()

    if key == "darkmode":
        # Chuẩn hóa về "0"/"1"
        norm = "1" if value.lower() in ("1", "true", "on", "yes") else "0"
        session["darkmode"] = norm
        return jsonify({"success": True})

    if key == "lang":
        # Giới hạn mã ngôn ngữ và độ dài
        norm = value.lower()[:5]
        if norm not in ("vi", "en"):
            norm = "vi"
        session["lang"] = norm
        return jsonify({"success": True})

    return jsonify({"success": False}), 400

KIOSK_TOKEN = os.environ.get("KIOSK_TOKEN") or ("kiosk-" + secrets.token_urlsafe(18))

def _kiosk_ok(req):
    """Chỉ cho phép truy cập khi có token đúng (đặt ?t=<token> trên URL)."""
    return req.args.get("t") == KIOSK_TOKEN

# ------------------------------------------------------------------------------
# 2) Bộ nhớ đệm dữ liệu cuối cùng của trang home.html (để cấp cho kiosk)
#    -> Không đụng vào hàm load dữ liệu của bạn, chỉ nghe lúc template render.
# ------------------------------------------------------------------------------
_last_home_ctx = {
    "summary_by_type": [],           # dạng: [{"short":"TR","late":1,"due":2,"must":0,"active":5,"total":8}, ...]
    "report_list": [],               # dạng: [{"report":"25-xxxx","item":"...","type_of":"...","status":"DUE","log_date":"YYYY-MM-DD","etd":"YYYY-MM-DD"}, ...]
    "counter": {"office": 0, "ot": 0},  # {"office": <HC done>, "ot": <OT done>}
    "generated_at": None
}

def _extract_for_kiosk(context: dict):
    """
    Từ context render của home.html, rút gọn dữ liệu cần cho kiosk.
    Hàm này an toàn nếu thiếu biến (sẽ dùng default).
    """
    summary_by_type = context.get("summary_by_type") or []
    report_list     = context.get("report_list") or []
    counter         = context.get("counter") or {"office": 0, "ot": 0}

    # Chuẩn hoá từng phần tử để đảm bảo key đầy đủ
    def _norm_summary(x):
        return {
            "short":  x.get("short", ""),
            "late":   int(x.get("late", 0) or 0),
            "due":    int(x.get("due", 0) or 0),
            "must":   int(x.get("must", 0) or 0),
            "active": int(x.get("active", 0) or 0),
            "total":  int(x.get("total", 0) or 0),
        }

    def _norm_report(r):
        return {
            "report":   r.get("report", "") or "",
            "item":     r.get("item", "") or "",
            "type_of":  r.get("type_of", "") or "",
            "status":   r.get("status", "") or "",
            "log_date": r.get("log_date", "") or "",
            "etd":      r.get("etd", "") or "",
        }

    norm_summary = [_norm_summary(x) for x in summary_by_type if isinstance(x, dict)]
    norm_reports = [_norm_report(r)  for r in report_list      if isinstance(r, dict)]
    norm_counter = {
        "office": int((counter or {}).get("office", 0) or 0),
        "ot":     int((counter or {}).get("ot", 0) or 0),
    }

    return {
        "summary_by_type": norm_summary,
        "report_list": norm_reports,
        "counter": norm_counter,
        "generated_at": datetime.now().isoformat(timespec="seconds")
    }

@template_rendered.connect_via(app)
def _capture_home_context(sender, template, context, **extra):
    """
    Nghe signal mỗi khi Flask render template nào đó.
    Khi template là 'home.html', lấy những biến cần và cache vào _last_home_ctx.
    """
    try:
        # Nếu tên file template của trang chính không phải 'home.html', đổi lại tại đây:
        if getattr(template, "name", None) == "home.html":
            data = _extract_for_kiosk(context or {})
            _last_home_ctx.update(copy.deepcopy(data))
    except Exception:
        # không để lỗi tại đây phá render của app
        pass

# ------------------------------------------------------------------------------
# 3) API dữ liệu cho kiosk: /api/display_data?t=<KIOSK_TOKEN>
# ------------------------------------------------------------------------------
@app.route("/api/display_data")
def api_display_data():
    if not _kiosk_ok(request):
        abort(403)

    # Nếu muốn fallback (khi server mới khởi động, chưa render home lần nào),
    # bạn có thể tự gọi hàm load dữ liệu của bạn ở đây, ví dụ:
    #
    # try:
    #     from yourmodule import load_home_data    # nếu bạn có sẵn hàm này
    #     summary_by_type, report_list, counter = load_home_data()
    #     data = _extract_for_kiosk({
    #         "summary_by_type": summary_by_type,
    #         "report_list": report_list,
    #         "counter": counter
    #     })
    #     _last_home_ctx.update(copy.deepcopy(data))
    # except Exception:
    #     pass
    #
    # Mặc định sẽ trả về cache gần nhất đã bắt được khi render home.html

    return jsonify({
        "generated_at": _last_home_ctx.get("generated_at") or datetime.now().isoformat(timespec="seconds"),
        "summary": _last_home_ctx.get("summary_by_type", []),
        "reports": _last_home_ctx.get("report_list", []),
        "counter": _last_home_ctx.get("counter", {"office": 0, "ot": 0})
    })

# ------------------------------------------------------------------------------
# 4) Trang kiosk: /display?t=<KIOSK_TOKEN>&page_len=15&rotate_sec=60&refresh_sec=30&dark=1
# ------------------------------------------------------------------------------
@app.route("/display")
def display_board():
    if not _kiosk_ok(request):
        abort(403)

    # Tham số cấu hình nhanh qua URL
    page_len    = int(request.args.get("page_len", 15))     # số dòng mỗi trang chi tiết
    rotate_sec  = int(request.args.get("rotate_sec", 60))   # lật trang mỗi X giây
    refresh_sec = int(request.args.get("refresh_sec", 30))  # nạp lại dữ liệu mỗi X giây
    dark        = request.args.get("dark", "1").lower() in ("1", "true", "yes")

    return render_template(
        "display.html",
        token=KIOSK_TOKEN,
        page_len=page_len,
        rotate_sec=rotate_sec,
        refresh_sec=refresh_sec,
        dark=dark
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8246,debug=True)
