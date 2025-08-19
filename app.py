from flask import Flask, request, render_template, session, redirect, url_for, jsonify, flash, send_from_directory, Response, stream_with_context, abort
from config import SECRET_KEY, local_main, SAMPLE_STORAGE, UPLOAD_FOLDER, TEST_GROUPS, local_complete, qr_folder, SO_GIO_TEST, ALL_SLOTS, TEAMS_WEBHOOK_URL_TRF, TEAMS_WEBHOOK_URL_RATE, TEAMS_WEBHOOK_URL_COUNT
from excel_utils import get_item_code, get_col_idx, copy_row_with_style, write_tfr_to_excel, append_row_to_trf
from image_utils import allowed_file, get_img_urls
from auth import login, get_user_type
from test_logic import load_group_notes, get_group_test_status, is_drop_test, is_impact_test, is_rotational_test,  TEST_GROUP_TITLES, TEST_TYPE_VI, DROP_ZONES, DROP_LABELS, GT68_FACE_LABELS, GT68_FACE_ZONES
from test_logic import IMPACT_ZONES, IMPACT_LABELS, ROT_LABELS, ROT_ZONES, RH_IMPACT_ZONES, RH_VIB_ZONES, RH_SECOND_IMPACT_ZONES, RH_STEP12_ZONES, update_group_note_file, get_group_note_value, F2057_TEST_TITLES
from notify_utils import send_teams_message, notify_when_enough_time
from counter_utils import update_counter, check_and_reset_counter, log_report_complete
from docx_utils import approve_request_fill_docx_pdf
from file_utils import safe_write_json, safe_read_json, safe_save_excel, safe_load_excel, safe_write_text, safe_read_text
import re, os, pytz, json, openpyxl, random, subprocess, traceback, regex, calendar, time
from contextlib import contextmanager
from datetime import datetime
from waitress import serve
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from collections import defaultdict, OrderedDict
from apscheduler.schedulers.background import BackgroundScheduler
from threading import Lock
from werkzeug.utils import secure_filename
from vfr3 import vfr3_bp

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.register_blueprint(vfr3_bp)

# Những test dùng giao diện Hot & Cold
HOTCOLD_LIKE = {"hot_cold", "standing_water", "stain","corrosion"}
INDOOR_GROUPS = {"indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal","outdoor_finishing"}
REPORT_NO_LOCK = Lock()
BLANK_TOKENS = {"", "-", "—"}

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
def _require_staff():
    # Yêu cầu đã nhập Staff ID để vào VFR3
    if not session.get('staff_id'):
        abort(403)

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

def calculate_default_etd(request_date_str, test_group):
    from datetime import datetime, timedelta
    if not request_date_str:
        return ""
    try:
        dt = datetime.strptime(request_date_str, "%Y-%m-%d")
    except:
        try:
            dt = datetime.strptime(request_date_str, "%d/%m/%Y")
        except:
            return ""
    test_group_code = ""
    if test_group:
        test_group_code = test_group.upper()[:3]
    days_to_add = 0
    if test_group_code in ["CON", "TRA"]:
        days_to_add = 3
    elif test_group_code in ["MAT", "FIN"]:
        days_to_add = 5
    else:
        days_to_add = 0
    if days_to_add > 0:
        etd_date = dt + timedelta(days=days_to_add)
        return etd_date.strftime("%Y-%m-%d")
    return ""

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
    lock_path = "tfr_report.lock"
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
            break
        except FileExistsError:
            time.sleep(0.05 + random.random() * 0.15)
    try:
        yield
    finally:
        os.close(fd)
        try:
            os.remove(lock_path)
        except:
            pass

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
    - Không bump tuần hoàn theo 'mã có trong Excel' vì cột B luôn có sẵn toàn bộ mã.
    - Có retry nhẹ (2 lần) để chống race-condition hiếm gặp.
    """
    with report_lock():
        tries = 0

        # Case A: đã có report_no trong req -> validate & dùng đúng số này
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
            # xác nhận lại: dòng vẫn còn trống?
            if not row_is_filled_for_report(local_main, report_no):
                return pdf_path, report_no

            # hi hữu: ai đó vừa điền vào dòng này giữa chừng -> thử lại một lần
            tries += 1
            if tries >= max_try:
                raise RuntimeError("Không tìm được dòng trống để cấp mã report.")
            # xoá file vừa sinh (đi nhầm dòng)
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
            # ép số mới vào req để make_report_func dùng đúng số này
            fixed_req = dict(req)
            fixed_req["report_no"] = bumped
            pdf_path, report_no = make_report_func(fixed_req)

        return pdf_path, report_no

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

        from datetime import datetime
        import pytz

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
        late = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "LATE")
        due = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "DUE")
        must = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "MUST")
        active = sum(1 for r in full_report_list if r['type_of'] == t and r['status'] == "ACTIVE")
        total = late + due + must + active
        summary_by_type.append({
            "type_of": t,
            "short": type_shortname.get(t, t[:3].upper()),
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

TFR_LOG_FILE = "tfr_requests.json"  # Dùng file json cho đơn giản

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

@app.route("/tfr_request_form", methods=["GET", "POST"])
def tfr_request_form():
    tfr_requests = safe_read_json(TFR_LOG_FILE)
    error = ""
    form_data = {}
    missing_fields = []

    # Lấy tham số từ URL hoặc từ POST
    trq_id = request.args.get("trq_id") or request.form.get("trq_id")
    edit_idx = request.args.get("edit_idx") or request.form.get("edit_idx")
    editing = False

    # Nếu có trq_id + edit_idx -> đang ở chế độ EDIT: nạp sẵn dữ liệu vào form_data
    if trq_id is not None and edit_idx is not None:
        try:
            abs_idx = int(edit_idx)
            # Ưu tiên: coi edit_idx là index tuyệt đối
            if 0 <= abs_idx < len(tfr_requests) and tfr_requests[abs_idx].get("trq_id") == trq_id:
                form_data = tfr_requests[abs_idx].copy()
                editing = True
            else:
                # Fallback: logic cũ theo ordinal trong nhóm cùng trq_id
                matches = [i for i, req in enumerate(tfr_requests) if req.get("trq_id") == trq_id]
                if len(matches) > abs_idx:
                    form_data = tfr_requests[matches[abs_idx]].copy()
                    editing = True
        except Exception:
            pass

    # Giữ lại edit_idx trong form_data để template render hidden input
    if editing:
        form_data.setdefault("edit_idx", edit_idx)

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
            "request_date": form.get("request_date"),
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

        # Tự tính ETD nếu chưa có
        new_request["etd"] = calculate_default_etd(
            new_request.get("request_date", ""),
            new_request.get("test_group", "")
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

        # Ghi đè item cũ hoặc append mới
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
            tfr_requests.append(new_request)

        safe_write_json(TFR_LOG_FILE, tfr_requests)

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

    # GET lần đầu (không EDIT) -> auto fill employee_id, requestor từ session
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

    # Mặc định request_date = hôm nay nếu trống
    if not form_data.get("request_date"):
        vn_tz = pytz.timezone("Asia/Ho_Chi_Minh")
        today = datetime.now(vn_tz).strftime("%Y-%m-%d")
        form_data["request_date"] = today

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

# --- ADD NEW: gom logic approve 1 request để tái dùng ---
def approve_all_one(req):
    """
    Approve 1 request:
      - cấp report_no + tạo DOCX/PDF
      - cập nhật Excel + TRF.xlsx
      - đẩy vào archive (short_data)
      - trả về req đã cập nhật (status/report_no/pdf_path/docx_path)
    """
    with REPORT_NO_LOCK:
        # đọc mới nhất để tránh “đua”
        current_list = safe_read_json(TFR_LOG_FILE)
        pdf_path, report_no = allocate_unique_report_no(
            approve_request_fill_docx_pdf, req, current_list
        )

    req["status"] = "Approved"
    req["decline_reason"] = ""
    req["report_no"] = report_no

    output_folder = os.path.join('static', 'TFR')
    output_docx = os.path.join(output_folder, f"{report_no}.docx")
    output_pdf = os.path.join(output_folder, f"{report_no}.pdf")

    # FALLBACK PDF: nếu convert lỗi (pythoncom/pywin32), không crash — cho phép dùng DOCX
    try:
        # nếu bạn có hàm try_convert_to_pdf thì gọi ở đây, an toàn với try/except
        if not os.path.exists(output_pdf):
            from docx_utils import try_convert_to_pdf
            try_convert_to_pdf(output_docx, output_pdf)  # hàm này sẽ raise nếu thiếu pythoncom
    except Exception as _pdf_e:
        # ghi log nhẹ để debug, nhưng KHÔNG dừng approve
        print("PDF convert failed, fallback to DOCX:", _pdf_e)

    # gán đường dẫn ưu tiên PDF nếu đã tồn tại
    if os.path.exists(output_pdf):
        req['pdf_path'] = f"TFR/{report_no}.pdf"
        req['docx_path'] = None
    else:
        req['pdf_path'] = None
        req['docx_path'] = f"TFR/{report_no}.docx"

    # Ghi Excel (giữ đúng cột/format như code cũ)
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
                            cell.number_format = 'd-mmm'
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
            if etd_val:
                set_val("etd", format_excel_date_short(etd_val), is_date_col=True)
            else:
                set_val("etd", "")

            from datetime import datetime
            import pytz
            vn_tz = pytz.timezone("Asia/Ho_Chi_Minh")
            log_in_date = datetime.now(vn_tz)
            set_val("log in date", format_excel_date_short(log_in_date), is_date_col=True)

            finishing_type = req.get("finishing_type", "")
            material_type  = req.get("material_type", "")
            cat_comp_pos   = get_category_component_position(finishing_type, material_type)
            set_val("category / component name / position", cat_comp_pos)
            wb.save(local_main)
    except Exception as e:
        print("Ghi vào Excel bị lỗi:", e)

    # Cập nhật TRF.xlsx
    try:
        append_row_to_trf(report_no, local_main, "TRF.xlsx", trq_id=req.get("trq_id", ""))
    except Exception as e:
        print("Append TRF lỗi:", e)

    # Đưa vào archive (archive_request của bạn đang tự dọn >14 ngày theo request_date)
    try:
        from datetime import datetime
        import pytz
        short_data = {
            "trq_id": req.get("trq_id", ""),
            "report_no": req.get("report_no", ""),
            "requestor": req.get("requestor", ""),
            "department": req.get("department", ""),
            "request_date": req.get("request_date", ""),
            "status": req.get("status", ""),
            "pdf_path": req.get("pdf_path"),
            "docx_path": req.get("docx_path"),
            "approved_date": datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")).strftime("%Y-%m-%d"),
        }
        archive_request(short_data)
    except Exception as e:
        print("Archive lỗi:", e)

    return req

@app.post("/approve_all_stream")
def approve_all_stream():
    def gen():
        # (0) Nhận payload
        try:
            data = request.get_json(silent=True) or {}
            updates = data.get("updates", []) or []
            run_id  = (data.get("run_id") or "").strip()
        except Exception as e:
            yield json.dumps({"type": "error", "message": f"Parse JSON: {e}"}) + "\n"
            return

        # Không có gì để duyệt → kết thúc sớm
        if not updates:
            yield json.dumps({"type": "start", "total": 0}) + "\n"
            yield json.dumps({"type": "done", "done": 0, "total": 0}) + "\n"
            return

        # (1) Cleanup archive (không chặn tiến trình)
        try:
            cleanup_archive_json(days=14)
        except Exception as e:
            yield json.dumps({"type": "error", "message": f"Cleanup: {e}"}) + "\n"

        # (2) Đọc pending snapshot (LIST thuần) + map id->index
        with PENDING_LOCK:
            pending_snapshot = safe_read_json(TFR_LOG_FILE)

        def make_id_index_map(lst):
            mp = {}
            for i, r in enumerate(lst):
                tid = (r.get("trq_id") or "").strip()
                if tid:  # chỉ map những dòng có trq_id hợp lệ
                    mp[tid] = i
            return mp

        id_to_idx = make_id_index_map(pending_snapshot)

        # (3) Cập nhật ETD CHỈ cho các item nằm trong 'updates' (theo trq_id) TRÊN LIST
        try:
            changed = False
            for u in updates:
                tid = (u.get("trq_id") or "").strip()
                etd = (u.get("etd") or "").strip()
                if tid and tid in id_to_idx:
                    pending_snapshot[id_to_idx[tid]]["etd"] = etd
                    changed = True
            if changed:
                with PENDING_LOCK:
                    safe_write_json(TFR_LOG_FILE, pending_snapshot)
                # đồng bộ lại
                with PENDING_LOCK:
                    pending_snapshot = safe_read_json(TFR_LOG_FILE)
                id_to_idx = make_id_index_map(pending_snapshot)
        except Exception as e:
            yield json.dumps({"type": "error", "message": f"Bulk ETD update: {e}"}) + "\n"

        # (4) Lập TODO theo đúng THỨ TỰ 'updates' (chỉ Submitted + có ETD)
        todo = []
        seen = set()
        for u in updates:
            tid = (u.get("trq_id") or "").strip()
            if not tid or tid in seen:
                continue
            seen.add(tid)
            idx = id_to_idx.get(tid)
            if idx is None:
                continue
            item = pending_snapshot[idx]
            if item and item.get("status") == "Submitted" and (item.get("etd") or "").strip():
                # chụp bản sao nhỏ để approve nhưng vẫn giữ pending_snapshot là nguồn sự thật
                todo.append((tid, item))

        total = len(todo)
        yield json.dumps({"type": "start", "total": total}) + "\n"
        if total == 0:
            yield json.dumps({"type": "done", "done": 0, "total": 0}) + "\n"
            return

        # (5) Vòng duyệt CHỈ trên 'todo', có kiểm tra hủy trước MỖI ITEM
        done = 0
        # Làm việc trên LIST hiện tại; chỉ loại bỏ CHÍNH XÁC item đã approve
        current_list = list(pending_snapshot)

        for tid, item in todo:
            # Bị hủy?
            if run_id and run_id in APPROVE_CANCEL:
                # bỏ cờ và kết thúc ngay, KHÔNG ghi thêm thay đổi nào
                APPROVE_CANCEL.discard(run_id)
                yield json.dumps({"type": "done", "done": done, "total": total}) + "\n"
                return

            try:
                # Pipeline chuẩn 1 item
                approved = approve_all_one(item)

                # Đồng bộ trường hiển thị (không bắt buộc ghi vào file pending vì sắp loại bỏ)
                report_no = (approved or {}).get("report_no") or item.get("report_no")
                item.update({
                    "status": "Approved",
                    "decline_reason": "",
                    "report_no": report_no,
                    "pdf_path": (approved or {}).get("pdf_path"),
                    "docx_path": (approved or {}).get("docx_path"),
                })

                # Chỉ xóa đúng item đã approve khỏi pending và lưu file
                # (lọc theo trq_id; KHÔNG động vào dòng khác)
                new_list = []
                for r in current_list:
                    rid = (r.get("trq_id") or "").strip()
                    if rid != tid:
                        new_list.append(r)
                current_list = new_list

                with PENDING_LOCK:
                    safe_write_json(TFR_LOG_FILE, current_list)

                done += 1
                yield json.dumps({
                    "type": "progress",
                    "done": done,
                    "total": total,
                    "trq_id": tid,
                    "report_no": report_no
                }) + "\n"

            except Exception as e:
                # không dừng toàn bộ – báo lỗi item này thôi (KHÔNG xóa gì)
                yield json.dumps({"type": "error", "message": str(e), "trq_id": tid}) + "\n"

        # (6) Kết thúc vòng duyệt
        yield json.dumps({"type": "done", "done": done, "total": total}) + "\n"

    headers = {
        "Content-Type": "application/x-ndjson; charset=utf-8",
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",
    }
    return Response(stream_with_context(gen()), headers=headers)

APPROVE_CANCEL = set()
PENDING_LOCK = Lock()

# ==== [Endpoint hủy từ FE] ====
@app.post("/approve_all_cancel")
def approve_all_cancel():
    data = request.get_json(silent=True) or {}
    rid = (data.get("run_id") or "").strip()
    if rid:
        APPROVE_CANCEL.add(rid)
    return {"ok": True}

@app.route("/tfr_request_status", methods=["GET", "POST"])
def tfr_request_status():
    tfr_requests = safe_read_json(TFR_LOG_FILE)
    is_admin = session.get("user_type") in ("stl", "superadmin")

    if request.method == "POST":
        action = request.form.get("action")
        tfr_requests = safe_read_json(TFR_LOG_FILE)  # reload

        # === APPROVE ALL ===
        if is_admin and action == "approve_all":
            approved_count = 0
            # duyệt bản copy để có thể remove phần tử trong danh sách gốc
            current = safe_read_json(TFR_LOG_FILE)
            for idx, req in enumerate(current[:]):
                if req.get("status") != "Submitted":
                    continue
                etd = request.form.get(f"etd-{idx}", "").strip()
                if not etd:
                    continue

                # cập nhật ETD trước khi approve
                req["etd"] = etd
                req["estimated_completion_date"] = etd

                try:
                    # gom toàn bộ logic approve vào 1 hàm (giữ đúng cách ghi excel/trf/archive)
                    approve_all_one(req)
                    # XÓA request đã approve khỏi JSON pending để đỡ nặng
                    trq_id = req.get("trq_id")
                    current = [r for r in current if r.get("trq_id") != trq_id]
                    approved_count += 1
                except Exception as e:
                    print("Approve one (approve_all) error:", e)

            safe_write_json(TFR_LOG_FILE, current)
            flash(f"Đã duyệt {approved_count} request (chỉ duyệt các dòng đã có ETD)!")
            return redirect(url_for('tfr_request_status'))
        
        # === APPROVE SINGLE ===
        elif is_admin and action == "approve":
            trq_id = request.form.get("trq_id")
            edit_idx = int(request.form.get("edit_idx", 0)) if "edit_idx" in request.form else None
            matches = [i for i, req in enumerate(tfr_requests) if req.get("trq_id") == trq_id]
            idx = matches[edit_idx] if edit_idx is not None and edit_idx < len(matches) else (
                matches[0] if matches else None
            )
            if idx is not None:
                req = tfr_requests[idx]
                etd = request.form.get("etd", "").strip()
                if not etd:
                    flash("Bạn cần điền Estimated Completion Date (ETD) trước khi approve!")
                    return redirect(url_for('tfr_request_status'))

                req["etd"] = etd
                req["estimated_completion_date"] = etd

                try:
                    approve_all_one(req)      # gom full logic approve 1 chỗ
                    del tfr_requests[idx]     # XÓA ngay khỏi pending để giảm nặng file
                    safe_write_json(TFR_LOG_FILE, tfr_requests)
                except Exception as e:
                    print("Approve one (single) error:", e)
                    flash("Có lỗi khi approve, vui lòng thử lại.")
            return redirect(url_for('tfr_request_status'))

        # === DECLINE ===
        elif is_admin and action == "decline":
            trq_id = request.form.get("trq_id")
            reason = request.form.get("decline_reason", "").strip()
            edit_idx = int(request.form.get("edit_idx", 0)) if "edit_idx" in request.form else None
            matches = [i for i, req in enumerate(tfr_requests) if req.get("trq_id") == trq_id]
            idx = matches[edit_idx] if edit_idx is not None and edit_idx < len(matches) else (
                matches[0] if matches else None
            )
            if idx is not None:
                tfr_requests[idx]["status"] = "Declined"
                tfr_requests[idx]["decline_reason"] = reason
            safe_write_json(TFR_LOG_FILE, tfr_requests)
            return redirect(url_for('tfr_request_status'))

        # === DUPLICATE ===
        elif action == "duplicate":
            trq_id = request.form.get("trq_id")
            edit_idx = int(request.form.get("edit_idx", 0)) if "edit_idx" in request.form else None
            matches = [i for i, req in enumerate(tfr_requests) if req.get("trq_id") == trq_id]
            idx = matches[edit_idx] if edit_idx is not None and edit_idx < len(matches) else (
                matches[0] if matches else None
            )
            if idx is not None:
                old_req = tfr_requests[idx]
                new_req = old_req.copy()
                new_req["report_no"] = ""
                new_req["status"] = "Submitted"
                new_req["pdf_path"] = ""
                new_req["decline_reason"] = ""
                # Tự động điền etd khi duplicate
                new_req["etd"] = calculate_default_etd(new_req.get("request_date", ""), new_req.get("test_group", ""))
                tfr_requests.insert(idx + 1, new_req)
            safe_write_json(TFR_LOG_FILE, tfr_requests)
            return redirect(url_for('tfr_request_status'))

        # === DELETE ===
        elif action == "delete":
            trq_id = request.form.get("trq_id")
            edit_idx = request.form.get("edit_idx")
            if edit_idx is not None:
                try:
                    edit_idx = int(edit_idx)
                    deleted_req = tfr_requests.pop(edit_idx)
                    from notify_utils import send_teams_message
                    send_teams_message(
                        TEAMS_WEBHOOK_URL_TRF,
                        f"🗑️ [TRF] Đã có yêu cầu bị xóa!\n- TRQ-ID: {deleted_req.get('trq_id')}\n- Người thao tác: {session.get('staff_id', 'Không rõ')}"
                    )
                except Exception as e:
                    print("Xóa bị lỗi:", e)
            else:
                # fallback: xóa theo trq_id (trường hợp cũ)
                for i, req in enumerate(tfr_requests):
                    if req.get("trq_id") == trq_id:
                        deleted_req = tfr_requests.pop(i)
                        from notify_utils import send_teams_message
                        send_teams_message(
                            TEAMS_WEBHOOK_URL_TRF,
                            f"🗑️ [TRF] Đã có yêu cầu bị xóa!\n- TRQ-ID: {deleted_req.get('trq_id')}\n- Người thao tác: {session.get('staff_id', 'Không rõ')}"
                        )
                        break
            safe_write_json(TFR_LOG_FILE, tfr_requests)
            return redirect(url_for('tfr_request_status'))

    # ===== FILTER SAU CÙNG: chỉ show request chưa duyệt =====
    show_requests = []
    real_indices = []
    for i, r in enumerate(tfr_requests):
        if r.get("status") == "Submitted":
            show_requests.append(r)
            real_indices.append(i)
    return render_template(
        "tfr_request_status.html",
        requests=show_requests,
        is_admin=is_admin,
        real_indices=real_indices
    )

@app.route("/tfr_request_archive")
def tfr_request_archive():

    archive = safe_read_json(ARCHIVE_LOG)
    now = datetime.now()

    def get_dt(d):
        if "-" in d: 
            return datetime.strptime(d, "%Y-%m-%d")
        else: 
            return datetime.strptime(d, "%d/%m/%Y")

    def safe_report_no(val):
        if isinstance(val, int):
            return val
        if isinstance(val, str):
            # Tách lấy phần sau dấu '-'
            if "-" in val:
                parts = val.split("-")
                try:
                    return int(parts[-1])
                except:
                    pass
            # Nếu không thì thử chuyển trực tiếp
            try:
                return int(val)
            except:
                return 0
        return 0

    archive = sorted(
        archive,
        key=lambda r: (
            get_dt(r["request_date"]), 
            safe_report_no(r.get("report_no", 0))
        ),
        reverse=True
    )

    return render_template("tfr_request_archive.html", requests=archive)

@app.route('/save_etd', methods=['POST'])
def save_etd():
    if not request.is_json:
        return jsonify(success=False, message="Invalid request"), 400
    data = request.get_json()
    trq_id = data.get("trq_id")
    etd = data.get("etd")
    idx = data.get("idx")
    if trq_id is None or etd is None or idx is None:
        return jsonify(success=False, message="Thiếu thông tin"), 400

    tfr_requests = safe_read_json(TFR_LOG_FILE)
    try:
        idx = int(idx)
        if idx < 0 or idx >= len(tfr_requests):
            return jsonify(success=False, message="Sai index!"), 404
        # Cập nhật trực tiếp vào dòng idx, KHÔNG CẦN DÒ THEO trq_id NỮA!
        tfr_requests[idx]["etd"] = etd
        safe_write_json(TFR_LOG_FILE, tfr_requests)
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
        script_path = r"C:\VFR\lab_update_app\excel export.py"
        result = subprocess.run([python_path, script_path],
                                shell=False, capture_output=True, text=True, timeout=900)
        if result.returncode == 0:
            return jsonify({'success': True, 'message': 'Đã chạy xong export file Excel!', 'reload': True})
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
    session.pop("auth_ok", None)
    session.pop("staff_id", None)  # Đăng xuất thì xóa luôn staff_id
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

    try:
        wb = safe_load_excel(local_main)
        ws = wb.active
        report_col = get_col_idx(ws, "report#") or get_col_idx(ws, "report")
        if report_col is None:
            return "❌ Không tìm thấy cột REPORT# hoặc REPORT trong file Excel!", 500

        # Tìm dòng theo report
        for row in range(2, ws.max_row + 1):
            v = ws.cell(row=row, column=report_col).value
            if v and str(v).strip() == str(report):
                row_idx = row
                break
        if row_idx is None:
            return f"❌ Không tìm thấy mã report {report} trong file Excel!", 404

        valid = True

# LẤY DATA CHO HIỂN THỊ (info_line)
        if not is_logged_in:
            summary_keys = [
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
            for col in range(2, ws.max_column + 1):
                label = ws.cell(row=1, column=col).value
                value = ws.cell(row=row_idx, column=col).value
                if label and value not in (None, ""):
                    lines.append((str(label).upper(), str(value)))
    except Exception as e:
        print("Lỗi khi đọc file excel:", e)
        print(traceback.format_exc())
        return f"Lỗi khi xử lý file: {e}", 500

    # --- XỬ LÝ LOGIN (nếu chưa đăng nhập) ---
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
            test_groups=TEST_GROUPS
        )

    # === ĐÃ ĐĂNG NHẬP: XỬ LÝ POST ===
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
            ws.cell(row=row_idx, column=complete_col).value = now
            safe_save_excel(wb, local_main)
            message = f"Đã ghi hoàn thành test cho {report}!"

        elif valid and action and action.startswith("rating_"):
            print("==> ĐANG XỬ LÝ RATING:", action, "CHO REPORT", report)
            value = action.replace("rating_", "").upper()

            # DÙNG SAFE LOAD để tránh xung đột file Excel
            wb = safe_load_excel(local_main)
            ws = wb.active

            rating_col = get_col_idx(ws, "rating")
            status_col = get_col_idx(ws, "status")
            ws.cell(row=row_idx, column=rating_col).value = value

            # --- LẤY LOẠI TEST GẦN NHẤT (từ session hoặc từ type_of Excel) ---
            last_test_type = session.get(f"last_test_type_{report}")
            type_of = ""
            group_code = None
            group_title = None
            if last_test_type:
                group_title = last_test_type
                for g_id, g_name in TEST_GROUPS:
                    if g_name == last_test_type:
                        group_code = g_id
                        break
            if not group_code:
                type_of_col = get_col_idx(ws, "type of")
                if type_of_col:
                    type_of = ws.cell(row=row_idx, column=type_of_col).value or ""
                group_code = str(type_of).strip().lower().replace(" ", "_")
                group_title = get_group_title(group_code) or (type_of or "")

            country_col = get_col_idx(ws, "country of destination")
            furniture_testing_col = get_col_idx(ws, "furniture testing")
            country = ws.cell(row=row_idx, column=country_col).value if country_col else ""
            furniture_testing = ws.cell(row=row_idx, column=furniture_testing_col).value if furniture_testing_col else ""

            # ======= Lấy thêm các trường bổ sung =======
            item_col = get_col_idx(ws, "item#")
            item_code_col = get_col_idx(ws, "item code") or get_col_idx(ws, "item_code")
            desc_col = get_col_idx(ws, "item name/ description")
            requestor_col = get_col_idx(ws, "submiter in") or get_col_idx(ws, "submitter in charge") or get_col_idx(ws, "requestor")

            item = ws.cell(row=row_idx, column=item_col).value if item_col else ""
            item_code = ws.cell(row=row_idx, column=item_code_col).value if item_code_col else ""
            desc = ws.cell(row=row_idx, column=desc_col).value if desc_col else ""
            requestor = ws.cell(row=row_idx, column=requestor_col).value if requestor_col else ""

            # ======= ĐƯỜNG LINK detail tới mã report này =======
            report_url = f"{request.url_root.rstrip('/')}/update?report={report}"
            staff_id = session.get("staff_id", "Không rõ")

            # --- Chuẩn bị thông báo Teams ---
            teams_msg = None
            if value == "PASS":
                teams_msg = (
                    f"✅ **PASS** {group_title}\n"
                    f"- Report#: {report}\n"
                    f"- Item#: {item}\n"
                    f"- Item code: {item_code}\n"
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
                        f"- Report#: {report}\n"
                        f"- Item#: {item}\n"
                        f"- Item code: {item_code}\n"
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
                        f"- Report#: {report}\n"
                        f"- Item#: {item}\n"
                        f"- Item code: {item_code}\n"
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

            # --- Đánh dấu hoàn thành trên file ---
            if status_col:
                ws.cell(row=row_idx, column=status_col).value = "COMPLETE"
                fill_complete = PatternFill("solid", fgColor="BFBFBF")
                for col in range(2, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col).fill = fill_complete

            # --- Copy sang completed file ---
            # Dùng safe_load_excel + safe_save_excel để không race condition
            if os.path.exists(local_complete):
                wb_c = safe_load_excel(local_complete)
                ws_c = wb_c.active
            else:
                wb_c = Workbook()
                ws_c = wb_c.active
                # Copy header (dòng 1) cả value + style + width + height từ ws (file ds)
                for col in range(1, ws.max_column + 1):
                    from_cell = ws.cell(row=1, column=col)
                    to_cell = ws_c.cell(row=1, column=col)
                    to_cell.value = from_cell.value
                    if from_cell.font:
                        to_cell.font = from_cell.font.copy()
                    if from_cell.border:
                        to_cell.border = from_cell.border.copy()
                    if from_cell.fill:
                        to_cell.fill = from_cell.fill.copy()
                    if from_cell.protection:
                        to_cell.protection = from_cell.protection.copy()
                    if from_cell.alignment:
                        to_cell.alignment = from_cell.alignment.copy()
                    to_cell.number_format = from_cell.number_format
                    col_letter = from_cell.column_letter
                    ws_c.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
                ws_c.row_dimensions[1].height = ws.row_dimensions[1].height
                safe_save_excel(wb_c, local_complete)

            # --- Sửa CHỐT: luôn kiểm tra cột mã report ---
            report_idx_in_c = get_col_idx(ws_c, "report#")
            if report_idx_in_c is None:
                report_idx_in_c = get_col_idx(ws_c, "report")
            if report_idx_in_c is None:
                report_idx_in_c = 2  # fallback về cột 1 (A)

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

            # ==== PHẦN BỔ SUNG: Ghi log ngay khi hoàn thành ====
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
            # Lấy employee_id từ session
            employee_id = session.get("staff_id", "")
            log_report_complete(report, type_of, ca, employee_id)  # Ghi cả ID người thao tác
            # ==== HẾT PHẦN BỔ SUNG ====

            message = f"Đã cập nhật đánh giá: <b>{value}</b> cho {report}!"
            check_and_reset_counter()
            update_counter()

    # === Lấy loại test gần nhất (last_test_type) ===
    last_test_type = session.get(f"last_test_type_{report}")

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

    # === Hiện thông báo nếu đủ giờ và chưa có ảnh after ===
    show_line_test_done_notice = show_line_test_done and not has_after_img

    # === Trả về template ===
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
    )

@app.route("/test_group/<report>/<group>", methods=["GET", "POST"])
def test_group_page(report, group): # Import trong hàm để tránh circular import nếu cần

    session[f"last_test_type_{report}"] = get_group_title(group)
    group_titles = TEST_GROUP_TITLES.get(group)
    if not group_titles:
        return "Nhóm kiểm tra không tồn tại!", 404

    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    status_file = os.path.join(report_folder, f"status_{group}.txt")
    comment_file = os.path.join(report_folder, f"comment_{group}.txt")

    # Đọc toàn bộ status/comment cho group đó (file lưu dạng "Mục xx: PASS/FAIL/N/A")
    all_status = load_group_notes(status_file)
    all_comment = load_group_notes(comment_file)

    # Duyệt từng test_key để lấy trạng thái, comment và có ảnh hay không
    test_status = {}
    for key in group_titles:
        if key == "hot_cold":
            # Đọc trạng thái hot_cold từ file riêng
            hotcold_status_file = os.path.join(report_folder, f"hotcold_status_{group}.txt")
            if os.path.exists(hotcold_status_file):
                hotcold_status = safe_read_text(hotcold_status_file).strip()
            else:
                hotcold_status = None
            st = hotcold_status
        else:
            st = all_status.get(key)
        cm = all_comment.get(key)
        has_img = False
        if os.path.exists(report_folder):
            has_img = any(
                allowed_file(f) and f.startswith(f"test_{group}_{key}_")
                for f in os.listdir(report_folder)
            )
        test_status[key] = {
            'status': st,
            'comment': cm,
            'has_img': has_img
        }

    # Nếu là tủ US thì cần status cho từng step f2057
    f2057_status = {}
    if group == 'tu_us':
        for fkey in F2057_TEST_TITLES:
            # Đọc status/comment/ảnh từng step con như bình thường
            s = get_group_test_status(report, group, fkey)
            f2057_status[fkey] = s

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
    session[f"last_test_type_{report}"] = get_group_title(group)

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
                if line.strip().startswith(f"Mục {key}:"):
                    return line.strip().split(":", 1)[1].strip()
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
        if 'status' in request.form and not group.startswith("transit"):
            status = request.form['status']
            update_group_note_file(status_file, key, status)
        # Ghi comment
        if 'save_comment' in request.form:
            comment = request.form.get('comment_input', '').strip()
            update_group_note_file(comment_file, key, comment)
        # Cập nhật loại kiểm tra gần nhất
        vi_name = TEST_TYPE_VI.get(group, group.upper())
        session[f"last_test_type_{report}"] = vi_name
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
        for f in sorted(os.listdir(report_folder)):
            if allowed_file(f) and f.startswith(f"test_{group}_{key}"):
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

# Cho phép URL có/không có test_key (mặc định là 'hot_cold' để không phá link cũ)
@app.route("/hot_cold_test/<report>/<group>", defaults={'test_key': 'hot_cold'}, methods=["GET", "POST"])
@app.route("/hot_cold_test/<report>/<group>/<test_key>", methods=["GET", "POST"])
def hot_cold_test(report, group, test_key):
    from_line = request.args.get("from_line")

    # ====== Lấy tên hiển thị đúng theo test_key ======
    # Ưu tiên lấy từ TEST_GROUP_TITLES; nếu không có thì prettify từ key
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

    session[f"last_test_type_{report}"] = f"{display_title} ({group.upper()})"

    # ====== Chuẩn bị đường dẫn/lưu trữ (tách theo test_key) ======
    vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(folder, exist_ok=True)

    # Prefix riêng cho từng test
    prefix           = f"{test_key}_{group}"
    status_file      = os.path.join(folder, f"{prefix}_status.txt")
    comment_file     = os.path.join(folder, f"{prefix}_comment.txt")
    before_tag       = f"{test_key}_before_{group}"
    after_tag        = f"{test_key}_after_{group}"
    before_time_file = os.path.join(folder, f"{prefix}_before_time.txt")
    after_time_file  = os.path.join(folder, f"{prefix}_after_time.txt")
    duration_file    = os.path.join(folder, f"{prefix}_duration.txt")  # <— giờ đếm ngược

    # ====== Xử lý POST ======
    if request.method == "POST":
        # 1) Cập nhật trạng thái
        if "status" in request.form:
            safe_write_text(status_file, request.form["status"])

        # 2) Lưu ghi chú
        if "save_comment" in request.form:
            safe_write_text(comment_file, request.form.get("comment_input", ""))

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
            # Nếu không còn ảnh trước/sau thì xoá file time tương ứng
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

        # 5) Cập nhật thời gian test (giờ) — cho phép người dùng chọn
        if "set_duration" in request.form:
            raw = (request.form.get("duration") or "").strip()
            try:
                dur = float(raw)
                if dur <= 0: raise ValueError
                safe_write_text(duration_file, str(dur))
                flash("Đã cập nhật thời gian test.", "success")
            except Exception:
                flash("Giá trị thời gian không hợp lệ.", "danger")

        # quay lại GET để tránh resubmit
        session[f"last_test_type_{report}"] = f"{display_title} ({group.upper()})"
        return redirect(request.url)

    # ====== Đọc dữ liệu để render ======
    status  = (safe_read_text(status_file) or "").strip()
    comment = (safe_read_text(comment_file) or "").strip()

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

    # Thời gian đếm ngược (giờ): đọc từ file; fallback SO_GIO_TEST nếu trống/lỗi
    raw_duration = safe_read_text(duration_file)
    try:
        so_gio_test = float(raw_duration) if raw_duration not in (None, "") else float(SO_GIO_TEST)
    except Exception:
        so_gio_test = 4.0  # fallback

    # ====== Render ======
    return render_template(
        "hot_cold_test.html",
        report=report,
        group=group,
        test_key=test_key,                                    # truyền xuống template để hiển thị tên/đặt field
        title={'short': display_title, 'full': display_title},
        status=status,
        comment=comment,
        imgs_mo_ta=imgs_mo_ta,
        imgs_before=imgs_before,
        imgs_after=imgs_after,
        before_upload_time=before_upload_time,
        after_upload_time=after_upload_time,
        so_gio_test=so_gio_test,                              # JS đếm ngược dùng biến này
        from_line=from_line,
        before_tag=before_tag,   # <— thêm
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
    session[f"last_test_type_{report}"] = "LINE TEST"
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
        session[f"last_test_type_{report}"] = "LINE TEST"
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

SAMPLE_STORAGE_FILE = "sample_storage.json"

@app.route("/store_sample", methods=["GET", "POST"])
def store_sample():
    report = request.args.get("report")
    item_code = get_item_code(report)
    auto_sample_name = f"{report} - {item_code}" if report and item_code else ""
    error_msg = ""

    # Đọc sample storage an toàn
    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    if not isinstance(SAMPLE_STORAGE, dict):
        SAMPLE_STORAGE = {}

    # Kiểm tra đã có mẫu lưu với report+item_code này chưa
    found_location = None
    for loc, info in SAMPLE_STORAGE.items():
        if info.get("report") == report and info.get("item_code") == item_code:
            found_location = loc
            break

    if found_location:
        # Đã có mẫu => chuyển sang trang info mẫu đó
        return redirect(url_for("sample_map", location_id=found_location))

    # Nếu chưa có thì xử lý như cũ
    if request.method == "POST":
        sample_name = request.form.get("sample_name")
        sample_type = request.form.get("sample_type")
        height = request.form.get("height")
        width = request.form.get("width")
        months = request.form.get("months")
        note = request.form.get("note")
        used_slots = set(SAMPLE_STORAGE.keys())

        # Lọc slot phù hợp
        if months == "3":
            possible_slots = [slot for slot in ALL_SLOTS if "-B" in slot]
        else:
            possible_slots = [slot for slot in ALL_SLOTS if "-A" in slot]
        free_slots = [slot for slot in possible_slots if slot not in used_slots]

        if not free_slots:
            return "<h3>Hết chỗ lưu mẫu phù hợp!</h3><a href='/'>Quay về</a>"
        location_id = free_slots[0]
        # --- Đọc lại (tránh ghi đè khi có nhiều người thao tác đồng thời) ---
        SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
        if not isinstance(SAMPLE_STORAGE, dict):
            SAMPLE_STORAGE = {}
        SAMPLE_STORAGE[location_id] = {
            'report': report,
            'item_code': item_code,
            'sample_name': sample_name,
            'sample_type': sample_type,
            'height': height,
            'width': width,
            'months': months,
            'note': note
        }
        safe_write_json(SAMPLE_STORAGE_FILE, SAMPLE_STORAGE)
        return redirect(url_for("sample_map", location_id=location_id))

    return render_template(
        "sample_form.html",
        report=report,
        item_code=item_code,
        auto_sample_name=auto_sample_name
    )

@app.route('/sample_map')
def sample_map():
    location_id = request.args.get('location_id')
    # Luôn đọc dữ liệu từ file, không dùng biến toàn cục
    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    if not isinstance(SAMPLE_STORAGE, dict):
        SAMPLE_STORAGE = {}

    sample = SAMPLE_STORAGE.get(location_id)
    if not sample:
        return "Không tìm thấy mẫu", 404

    report = sample['report']
    item_code = sample['item_code']

    return render_template(
        "sample_infor.html",
        info=sample,
        report_id=report,
        item_code=item_code,
        location_id=location_id
    )

@app.route("/list_samples", methods=["GET", "POST"])
def list_samples():
    # Luôn đọc file dữ liệu mẫu
    SAMPLE_STORAGE = safe_read_json(SAMPLE_STORAGE_FILE)
    if not isinstance(SAMPLE_STORAGE, dict):
        SAMPLE_STORAGE = {}

    if request.method == "POST":
        loc = request.form.get("loc")
        borrower = request.form.get("borrower")
        note = request.form.get("note")
        if loc in SAMPLE_STORAGE:
            SAMPLE_STORAGE[loc]['borrower'] = borrower
            SAMPLE_STORAGE[loc]['note'] = note
            # Ghi lại sau khi update
            safe_write_json(SAMPLE_STORAGE_FILE, SAMPLE_STORAGE)

    edit_loc = request.args.get("edit")
    report_id = request.args.get("report", "")
    item_code = ""

    table_rows = []
    for loc, info in SAMPLE_STORAGE.items():
        if not report_id and info.get('report', ''):
            report_id = info.get('report', '')
            item_code = info.get('item_code', '')
        table_rows.append({
            "loc": loc,
            "report": info.get('report', ''),
            "item_code": info.get('item_code', ''),
            "sample_type": info.get('sample_type', ''),
            "borrower": info.get('borrower', ''),
            "note": info.get('note', '')
        })

    return render_template(
        "list_samples.html",
        rows=table_rows,
        report_id=report_id,
        item_code=item_code,
        edit_loc=edit_loc
    )

@app.route('/images/<report>/imgs_<group>_<test_key>/<filename>')
def serve_test_img(report, group, test_key, filename):
    folder = os.path.join(UPLOAD_FOLDER, report, f"imgs_{group}_{test_key}")
    if not os.path.exists(folder):
        # Báo lỗi rõ ràng hoặc trả về 404
        return "Không tìm thấy thư mục ảnh!", 404
    return send_from_directory(folder, filename)

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

def auto_notify_all_first_time():
    webhook_url = TEAMS_WEBHOOK_URL_COUNT
    try:
        for report_folder in os.listdir(UPLOAD_FOLDER):
            folder = os.path.join(UPLOAD_FOLDER, report_folder)
            if not os.path.isdir(folder):
                continue

            # Line test: gửi ngay khi đủ giờ (job mỗi phút)
            try:
                notify_when_enough_time(
                    report=report_folder,
                    so_gio_test=SO_GIO_TEST,
                    tag_after="line_after",
                    time_file_name="before_upload_time.txt",
                    flag_file_name="teams_notified_line.txt",
                    webhook_url=webhook_url,
                    notify_msg=f"✅ [TỰ ĐỘNG] Line test của sản phẩm REPORT {report_folder} đã đủ {SO_GIO_TEST} tiếng! Vui lòng upload ảnh after.",
                    force_send=False,
                    pending_notify_name="pending_notify_line.txt"
                )
            except Exception as e:
                print(f"[auto_notify_all_first_time] Error notifying LINE for {report_folder}:", e)

            # Hotcold test: gửi ngay khi đủ giờ (job mỗi phút)
            for group in ["indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal"]:
                for key in HOTCOLD_LIKE:
                    try:
                        notify_when_enough_time(
                            report=report_folder,
                            so_gio_test=SO_GIO_TEST,
                            tag_after=f"{key}_after",                             # ví dụ: hot_cold_after
                            time_file_name=f"{key}_{group}_before_time.txt",      # ví dụ: hot_cold_indoor_thuong_before_time.txt
                            flag_file_name=f"teams_notified_{key}_{group}.txt",
                            webhook_url=webhook_url,
                            notify_msg=(f"✅ [TỰ ĐỘNG] {DISPLAY.get(key, key.title())} của REPORT {report_folder} "
                                        f"({group.upper()}) đã đủ {SO_GIO_TEST} tiếng! Vui lòng upload ảnh after."),
                            force_send=False,
                            pending_notify_name=f"pending_notify_{key}_{group}.txt"
                        )
                    except Exception as e:
                        print(f"[auto_notify_all_first_time] Error notifying {key} ({group}) for {report_folder}:", e)
    except Exception as e:
        print("[auto_notify_all_first_time] Error listing folders:", e)

def auto_notify_all_repeat():
    webhook_url = TEAMS_WEBHOOK_URL_COUNT
    MAX_REPEAT = 3

    def get_repeat_count(folder, file_name):
        path = os.path.join(folder, file_name)
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    val = f.read().strip()
                    return int(val) if val.isdigit() else 0
        except Exception:
            return 0
        return 0

    def increase_repeat_count(folder, file_name):
        path = os.path.join(folder, file_name)
        count = get_repeat_count(folder, file_name) + 1
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(str(count))
        except Exception:
            pass

    for report_folder in os.listdir(UPLOAD_FOLDER):
        folder = os.path.join(UPLOAD_FOLDER, report_folder)
        if not os.path.isdir(folder):
            continue

        # --- LINE TEST ---
        repeat_file_line = "repeat_notify_line.txt"
        count_line = get_repeat_count(folder, repeat_file_line)
        if count_line < MAX_REPEAT:
            sent = notify_when_enough_time(
                report=report_folder,
                so_gio_test=SO_GIO_TEST,
                tag_after="line_after",
                time_file_name="before_upload_time.txt",
                flag_file_name=None,
                webhook_url=webhook_url,
                notify_msg=f"✅ [TỰ ĐỘNG, NHẮC LẠI 1 TIẾNG] Line test của sản phẩm REPORT {report_folder} đã đủ {SO_GIO_TEST} tiếng! Vui lòng upload ảnh after.",
                force_send=True,
                pending_notify_name="pending_notify_line.txt"
            )
            # notify_when_enough_time nên trả về True nếu đã gửi notify lần này
            if sent:
                increase_repeat_count(folder, repeat_file_line)

        # --- HOTCOLD TEST ---
        for group in ["indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal"]:
            for key in HOTCOLD_LIKE:
                repeat_file = f"repeat_notify_{key}_{group}.txt"
                count = get_repeat_count(folder, repeat_file)
                if count < MAX_REPEAT:
                    sent = notify_when_enough_time(
                        report=report_folder,
                        so_gio_test=SO_GIO_TEST,
                        tag_after=f"{key}_after",
                        time_file_name=f"{key}_{group}_before_time.txt",
                        flag_file_name=None,
                        webhook_url=webhook_url,
                        notify_msg=(f"✅ [TỰ ĐỘNG, NHẮC LẠI 1 TIẾNG] {DISPLAY.get(key, key.title())} của REPORT "
                                    f"{report_folder} ({group.upper()}) đã đủ {SO_GIO_TEST} tiếng! Vui lòng upload ảnh after."),
                        force_send=True,
                        pending_notify_name=f"pending_notify_{key}_{group}.txt"
                    )
                    if sent:
                        increase_repeat_count(folder, repeat_file)

def auto_notify_all_pending():
    webhook_url = TEAMS_WEBHOOK_URL_COUNT
    # Luôn dùng giờ VN để không bị lệch khi server ở nước ngoài
    vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
    now = datetime.now(vn_tz)
    cur_hour = now.hour
    if cur_hour < 8 or cur_hour >= 21:
        return  # Chỉ gửi pending từ 8h tới 21h

    for report_folder in os.listdir(UPLOAD_FOLDER):
        folder = os.path.join(UPLOAD_FOLDER, report_folder)
        if not os.path.isdir(folder): continue

        # Line test
        pending_path = os.path.join(folder, "pending_notify_line.txt")
        if os.path.exists(pending_path):
            with open(pending_path, "r", encoding="utf-8") as f:
                msg = f.read()
            send_teams_message(webhook_url, msg)
            os.remove(pending_path)

        # Hotcold test
        for group in ["indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal"]:
            for key in HOTCOLD_LIKE:
                pending_path = os.path.join(folder, f"pending_notify_{key}_{group}.txt")
                if os.path.exists(pending_path):
                    with open(pending_path, "r", encoding="utf-8") as f:
                        msg = f.read()
                    send_teams_message(webhook_url, msg)
                    os.remove(pending_path)

# Khởi tạo scheduler
scheduler = BackgroundScheduler()
scheduler.add_job(func=auto_notify_all_first_time, trigger="interval", seconds=60)
scheduler.add_job(func=auto_notify_all_repeat, trigger="interval", seconds=3600)
scheduler.add_job(func=auto_notify_all_pending, trigger="interval", seconds=300)  # Kiểm tra pending mỗi 5 phút
scheduler.start()

@app.route("/set_pref", methods=["POST"])
def set_pref():
    key = request.json.get("key")
    value = request.json.get("value")
    if key in ("darkmode", "lang"):
        session[key] = value
        return jsonify({"success": True})
    return jsonify({"success": False}), 400

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8246,debug=True)
