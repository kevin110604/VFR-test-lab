# app.py
from flask import Flask, request, render_template, session, redirect, url_for
from config import SECRET_KEY, PASSWORD, local_main, SAMPLE_STORAGE, UPLOAD_FOLDER, TEST_GROUPS, local_complete, qr_folder, SO_GIO_TEST, ALL_SLOTS, TEAMS_WEBHOOK_URL
from excel_utils import get_item_code, get_col_idx, copy_row_with_style, is_img_at_cell, write_tfr_to_excel
from image_utils import allowed_file, safe_filename, get_img_urls
from auth import login, logout, is_logged_in
from test_logic import load_group_notes, get_group_test_status, is_drop_test, is_impact_test, is_rotational_test,  TEST_GROUP_TITLES, TEST_TYPE_VI, DROP_ZONES, DROP_LABELS
from test_logic import IMPACT_ZONES, IMPACT_LABELS, ROT_LABELS, ROT_ZONES, RH_IMPACT_ZONES, RH_VIB_ZONES, RH_SECOND_IMPACT_ZONES, RH_STEP12_ZONES, update_group_note_file, get_group_note_value
from notify_utils import send_teams_message, notify_when_enough_time
from counter_utils import update_counter, log_order_complete, check_and_reset_counter, log_report_complete
from openpyxl import load_workbook, Workbook
from flask import send_from_directory
from datetime import datetime
from openpyxl.styles import PatternFill
from excel_utils import ensure_column
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from apscheduler.schedulers.background import BackgroundScheduler
from docx_utils import approve_request_fill_docx_pdf
import re
import os
import pytz
import json
import openpyxl
import random
from collections import defaultdict, OrderedDict

app = Flask(__name__)
app.secret_key = SECRET_KEY

def get_group_title(group):
    for g_id, g_name in TEST_GROUPS:
        if g_id == group:
            return g_name
    return None

def generate_unique_tlq_id(existing_ids):
    while True:
        new_id = f"TL-{random.randint(100000, 999999)}"
        if new_id not in existing_ids:
            return new_id
        
@app.route("/", methods=["GET", "POST"])
def home():
    message = None

    # Đăng nhập
    if request.method == "POST":
        if request.form.get("action") == "login":
            if request.form.get("password") == PASSWORD:
                session["auth_ok"] = True
                return redirect(url_for("home"))
            else:
                message = "Sai mật khẩu. Vui lòng thử lại."

    # ==== Load danh sách report từ file ====
    full_report_list = []
    type_of_set = set()
    if session.get("auth_ok"):
        try:
            wb = load_workbook(local_main)
            ws = wb.active

            def clean_col(s):
                s = str(s).lower().strip()
                s = re.sub(r'[^a-z0-9#]+', '', s)
                return s

            headers = {}
            for col in range(1, ws.max_column + 1):
                name = ws.cell(row=1, column=col).value
                if name:
                    clean = clean_col(name)
                    headers[clean] = col

            report_col    = headers.get("report#")
            item_col      = headers.get("item#")
            status_col    = headers.get("status")
            test_date_col = headers.get("logindate")
            type_of_col   = headers.get("typeof")

            if None in (report_col, item_col, status_col, test_date_col):
                message = f"Thiếu cột trong file Excel! Đã đọc: {headers}"
            else:
                for row in range(2, ws.max_row + 1):
                    status_raw = ws.cell(row=row, column=status_col).value
                    status = str(status_raw).strip().upper() if status_raw else ""
                    if status not in ("LATE", "MUST", "DUE", "ACTIVE"):
                        continue

                    report = ws.cell(row=row, column=report_col).value
                    item = ws.cell(row=row, column=item_col).value
                    type_of = ws.cell(row=row, column=type_of_col).value if type_of_col else ""
                    log_date = ws.cell(row=row, column=test_date_col).value
                    log_date_str = str(log_date).strip() if log_date else ""
                    if log_date_str:
                        log_date_str = log_date_str.split()[0]

                    r_dict = {
                        "report": str(report).strip() if report else "",
                        "item": str(item).strip() if item else "",
                        "status": status,
                        "type_of": str(type_of).strip() if type_of else "",
                        "log_date": log_date_str
                    }
                    full_report_list.append(r_dict)
                    if r_dict["type_of"]:
                        type_of_set.add(r_dict["type_of"])

                type_of_set = sorted(type_of_set)
        except Exception as e:
            message = f"Lỗi khi đọc danh sách: {e}"

    # ==== Tổng hợp status toàn bộ ====
    type_shortname = {
        "CONSTRUCTION": "CON",
        "FINISHING": "FIN",
        "MATERIAL": "MAT",
        "PACKING": "PAC",
        "GENERAL": "GEN",
        # ... bổ sung nếu có loại khác
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

    # ==== Lọc report_list theo loại nếu có ====
    selected_type = request.args.get("type_of", "")
    if selected_type:
        report_list = [r for r in full_report_list if r["type_of"] == selected_type]
    else:
        report_list = full_report_list

    # Đọc số liệu đếm để truyền vào home.html
    counter = {"office": 0, "ot": 0}
    path = "counter_stats.json"
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            counter = json.load(f)

    # ==== Trả về template ====
    return render_template(
        "home.html",
        message=message,
        type_of_set=type_of_set,
        selected_type=selected_type,
        session=session,
        report_list=report_list,             # TRẢ VỀ FULL LIST SAU KHI LỌC!
        summary_by_type=summary_by_type,     # Vẫn có nếu template cần hiển thị tổng hợp
        counter=counter,
    )

TFR_LOG_FILE = "tfr_requests.json"  # Dùng file json cho đơn giản

@app.route("/tfr_request_form", methods=["GET", "POST"])
def tfr_request_form():
    # Đọc danh sách đã request (để hiển thị nếu cần)
    if os.path.exists(TFR_LOG_FILE):
        tfr_requests = []
        with open(TFR_LOG_FILE, "r", encoding="utf-8") as f:
            try:
                content = f.read().strip()
                if content:
                    tfr_requests = json.loads(content)
                else:
                    tfr_requests = []
            except Exception:
                tfr_requests = []
    else:
        tfr_requests = []

    error = ""
    form_data = {}
    missing_fields = []

    # Khi người dùng vừa vào form (GET), sinh sẵn TLQ-ID
    existing_ids = {r.get("tlq_id") for r in tfr_requests if "tlq_id" in r}
    default_tlq_id = generate_unique_tlq_id(existing_ids)

    if request.method == "POST":
        form = request.form

        # Validate các trường bắt buộc
        required_fields = [
            "requestor", "department", "request_date", "sample_type",
            "sample_description", "test_status", "quantity"
        ]
        for field in required_fields:
            if not form.get(field):
                missing_fields.append(field)

        # Validate nhóm test tối thiểu 1
        test_groups = request.form.getlist("test_group")
        if not test_groups:
            missing_fields.append("test_group")
            error = "Phải chọn ít nhất 1 loại test!"

        # Validate Furniture testing
        furniture_testing = request.form.getlist("furniture_testing")
        if not furniture_testing:
            missing_fields.append("furniture_testing")
            error = "Phải chọn Indoor hoặc Outdoor!"

        # Validate N/A logic cho các trường option
        def na_or_value(key):
            na_key = key + "_na"
            if form.get(na_key):
                return "N/A"
            return form.get(key, "").strip()

        form_data = form.to_dict(flat=False)  # Lưu lại tất cả đã nhập (kể cả multi checkbox)
        for k, v in form_data.items():
            if isinstance(v, list) and len(v) == 1:
                form_data[k] = v[0]
        form_data["test_group"] = test_groups
        form_data["furniture_testing"] = furniture_testing

        # Giữ lại ID để hiển thị khi có lỗi
        form_data["tlq_id"] = request.form.get("tlq_id", default_tlq_id)

        if missing_fields:
            if not error:
                error = "Vui lòng điền đủ các trường bắt buộc (*)"
            return render_template(
                "tfr_request_form.html",
                error=error,
                form_data=form_data,
                missing_fields=missing_fields
            )

        # Nếu không lỗi, tạo request mới
        item_code = na_or_value("item_code")
        supplier = na_or_value("supplier")
        subcon = na_or_value("subcon")
        test_status = form.get("test_status")
        if test_status == "nth":
            nth = form.get("test_status_nth", "").strip()
            test_status = nth + "th" if nth.isdigit() else "nth"
        furniture_testing_str = ", ".join(furniture_testing)
        new_request = {
            "tlq_id": form.get("tlq_id", default_tlq_id),
            "requestor": form.get("requestor"),
            "department": form.get("department"),
            "request_date": form.get("request_date"),
            "sample_type": form.get("sample_type"),
            "sample_description": form.get("sample_description"),
            "item_code": item_code,
            "supplier": supplier,
            "subcon": subcon,
            "test_status": test_status,
            "furniture_testing": furniture_testing_str,
            "quantity": form.get("quantity"),
            "test_groups": test_groups,
            "status": "Submitted",
            "decline_reason": "",
            "report_no": ""
        }
        tfr_requests.append(new_request)
        with open(TFR_LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(tfr_requests, f, ensure_ascii=False, indent=2)
        return redirect(url_for('tfr_request_status'))

    # GET – truy cập lần đầu, truyền sẵn ID vào form
    form_data["tlq_id"] = default_tlq_id
    return render_template("tfr_request_form.html", error=error, form_data=form_data, missing_fields=missing_fields)

@app.route("/tfr_request_status", methods=["GET", "POST"])
def tfr_request_status():
    # Đọc log request
    if os.path.exists(TFR_LOG_FILE):
        tfr_requests = []
        with open(TFR_LOG_FILE, "r", encoding="utf-8") as f:
            try:
                content = f.read().strip()
                if content:
                    tfr_requests = json.loads(content)
                else:
                    tfr_requests = []
            except Exception:
                tfr_requests = []
    else:
        tfr_requests = []

    is_admin = session.get("auth_ok", False)

    # Xử lý POST khi admin duyệt hoặc duplicate
    if request.method == "POST" and is_admin:
        idx = int(request.form.get("idx"))
        action = request.form.get("action")
        if 0 <= idx < len(tfr_requests):
            if action == "approve":
                etd_value = request.form.get("etd", "").strip()
                if not etd_value:
                    from flask import flash
                    flash("Bạn cần điền Estimated Completion Date (ETD) trước khi approve!")
                    return redirect(url_for('tfr_request_status'))
                tfr_requests[idx]["etd"] = etd_value  # Lưu ETD khi approve
                pdf_path, report_no = approve_request_fill_docx_pdf(tfr_requests[idx])
                tfr_requests[idx]["status"] = "Approved"
                tfr_requests[idx]["decline_reason"] = ""
                tfr_requests[idx]["pdf_path"] = pdf_path
                tfr_requests[idx]["report_no"] = report_no

                # == Ghi vào file Excel ==
                print("DEBUG: local_main =", local_main)
                write_tfr_to_excel(local_main, report_no, tfr_requests[idx])

                # === GHI THÔNG TIN VÀO EXCEL DS ===
                try:
                    # Load file Excel DS
                    wb = load_workbook(local_main)
                    ws = wb.active

                    # Tìm dòng có mã report đúng (so sánh với report_no vừa gán)
                    report_col = get_col_idx(ws, "report")
                    row_idx = None
                    for row in range(2, ws.max_row + 1):
                        v = ws.cell(row=row, column=report_col).value
                        if v and str(v).strip() == str(report_no):
                            row_idx = row
                            break
                    if row_idx:
                        def set_val(col_name, value):
                            col_idx = get_col_idx(ws, col_name)
                            if col_idx:
                                ws.cell(row=row_idx, column=col_idx).value = value

                        set_val("item#", tfr_requests[idx].get("item_code", ""))
                        # Nếu test_groups là list thì join lại, hoặc chỉ lấy 1 nhóm tùy bạn muốn
                        groups = tfr_requests[idx].get("test_groups", [])
                        if isinstance(groups, list):
                            groups_val = ", ".join(groups)
                        else:
                            groups_val = groups or ""
                        set_val("type of", groups_val)
                        set_val("item name/ description", tfr_requests[idx].get("sample_description", ""))
                        set_val("furniture testing", tfr_requests[idx].get("furniture_testing", ""))
                        set_val("submiter in", tfr_requests[idx].get("requestor", ""))
                        set_val("submited", tfr_requests[idx].get("department", ""))
                        set_val("remark", tfr_requests[idx].get("test_status", ""))
                        # Không đổi format file Excel, chỉ ghi value
                        wb.save(local_main)
                except Exception as e:
                    print("Ghi vào Excel bị lỗi:", e)

            elif action == "decline":
                tfr_requests[idx]["status"] = "Declined"
                tfr_requests[idx]["decline_reason"] = request.form.get("decline_reason", "")
            elif action == "delete":
                tfr_requests.pop(idx)
            elif action == "duplicate":
                old_req = tfr_requests[idx]
                new_req = old_req.copy()
                # Khi duplicate sẽ KHÔNG COPY report_no, chỉ clear nó để approve sẽ tự sinh
                new_req["report_no"] = ""
                new_req["status"] = "Submitted"
                new_req["pdf_path"] = ""
                new_req["decline_reason"] = ""
                new_req["etd"] = ""
                tfr_requests.append(new_req)

            # Không còn action save_etd nữa!
            with open(TFR_LOG_FILE, "w", encoding="utf-8") as f:
                json.dump(tfr_requests, f, ensure_ascii=False, indent=2)
        return redirect(url_for('tfr_request_status'))

    return render_template("tfr_request_status.html", requests=tfr_requests, is_admin=is_admin)

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
    img_path = os.path.join(UPLOAD_FOLDER, report, imgfile)   # ĐÚNG
    if os.path.exists(img_path):
        os.remove(img_path)
    return redirect(url_for('update', report=report))

@app.route("/delete_test_group_image/<report>/<group>/<key>/<imgfile>", methods=["POST"])
def delete_test_group_image(report, group, key, imgfile):
    img_path = os.path.join(UPLOAD_FOLDER, report, imgfile)
    if os.path.exists(img_path):
        os.remove(img_path)
    return redirect(url_for("test_group_item_dynamic", report=report, group=group, test_key=key))

@app.route("/logout")
def logout():
    session.pop("auth_ok", None)
    return "<h3 style='text-align:center;margin-top:80px;'>Đã đăng xuất!<br><a href='/' style='color:#4d665c;'>Về trang chọn sản phẩm</a></h3>"

@app.route("/update", methods=["GET", "POST"])
def update():
    report = request.args.get("report")
    if not report:
        return redirect("/")

    item_id, row_idx = None, None
    lines = []

    try:
        wb = load_workbook(local_main)
        ws = wb.active
        report_col = get_col_idx(ws, "report")
        row_idx = None

        for row in range(2, ws.max_row + 1):
            v = ws.cell(row=row, column=report_col).value
            if v and str(v).strip() == str(report):
                row_idx = row
                break

        if not session.get("auth_ok"):
            summary_keys = [
                ('report', 'REPORT#'),
                ('item', 'ITEM#'),
                ('type of', 'TYPE OF'),
                ('furniture testing', 'FURNITURE TESTING'),
                ('remark', 'REMARK'),
                ('qa comment', 'QA COMMENT'),
                ('rating', 'RATING')
            ]
            for key, label in summary_keys:
                idx_col = get_col_idx(ws, key)
                if idx_col:
                    value = ws.cell(row=row_idx, column=idx_col).value if row_idx else ""
                    if key == "rating":
                        show_value = str(value).strip() if value not in ("", None) else ""
                        lines.append((label, show_value))
                    else:
                        lines.append((label, value if value is not None else ""))
        else:
            if row_idx:
                for col in range(1, ws.max_column + 1):
                    label = ws.cell(row=1, column=col).value
                    value = ws.cell(row=row_idx, column=col).value
                    if label and value not in (None, ""):
                        lines.append((str(label).upper(), str(value)))
    except Exception:
        lines = []

    message = ""
    is_logged_in = session.get("auth_ok", False)
    valid = row_idx is not None

    if not is_logged_in:
        if request.method == "POST" and request.form.get("action") == "login":
            if request.form.get("password") == PASSWORD:
                session["auth_ok"] = True
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

    # Nếu đã đăng nhập và có thao tác POST
    if request.method == "POST":
        action = request.form.get("action")

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

        elif valid and action == "testing":
            wb = load_workbook(local_main)
            ws = wb.active
            test_date_col = get_col_idx(ws, "test date")
            vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
            now = datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M").upper()
            ws.cell(row=row_idx, column=test_date_col).value = now
            wb.save(local_main)
            message = f"Đã ghi thời gian kiểm tra cho {report}!"

        elif valid and action == "test_done":
            wb = load_workbook(local_main)
            ws = wb.active
            complete_col = get_col_idx(ws, "complete date")
            vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
            now = datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M").upper()
            ws.cell(row=row_idx, column=complete_col).value = now
            wb.save(local_main)
            message = f"Đã ghi hoàn thành test cho {report}!"

        elif valid and action and action.startswith("rating_"):
            print("==> ĐANG XỬ LÝ RATING:", action, "CHO REPORT", report)
            value = action.replace("rating_", "").upper()
            # workbook & worksheet đã được load đầu hàm, không cần load lại

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

            # --- Chuẩn bị thông báo Teams ---
            teams_msg = None
            if value == "PASS":
                teams_msg = (
                    f"✅ **PASS** {group_title}\n"
                    f"- Report#: {report}\n"
                    f"- Country of Destination: {country}\n"
                    f"- Furniture Testing: {furniture_testing}"
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
                        f"- Các mục FAIL:\n"
                        + "\n".join(group_fails)
                    )
                else:
                    teams_msg = (
                        f"{status_text} {group_title}\n"
                        f"- Report#: {report}\n"
                        f"- Không có mục nào FAIL trong nhóm này."
                    )
            if teams_msg:
                send_teams_message(TEAMS_WEBHOOK_URL, teams_msg)

            # --- Đánh dấu hoàn thành trên file ---
            if status_col:
                ws.cell(row=row_idx, column=status_col).value = "COMPLETE"
                fill_complete = PatternFill("solid", fgColor="BFBFBF")
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col).fill = fill_complete

            # --- Copy sang completed file ---
            if os.path.exists(local_complete):
                wb_c = load_workbook(local_complete)
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
                wb_c.save(local_complete)

            report_idx_in_c = get_col_idx(ws_c, "report")
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

            wb_c.save(local_complete)
            wb.save(local_main)

            # ==== PHẦN BỔ SUNG: Ghi log ngay khi hoàn thành ====
            from counter_utils import log_report_complete
            type_of_col = get_col_idx(ws, "type of")
            type_of = ws.cell(row=row_idx, column=type_of_col).value if type_of_col else ""
            # Xác định ca
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

            log_report_complete(report, type_of, ca)
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
def test_group_page(report, group):
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

    # Kiểm tra/tạo thư mục report_folder trước khi listdir
    if os.path.exists(report_folder):
        file_list = os.listdir(report_folder)
    else:
        file_list = []

    # Duyệt từng test_key để lấy trạng thái, comment và có ảnh hay không
    test_status = {}
    for key in group_titles:
        if key == "hot_cold":
            # Đọc trạng thái hot_cold từ file riêng
            hotcold_status_file = os.path.join(report_folder, f"hotcold_status_{group}.txt")
            if os.path.exists(hotcold_status_file):
                with open(hotcold_status_file, "r", encoding="utf-8") as f:
                    hotcold_status = f.read().strip()
            else:
                hotcold_status = None
            st = hotcold_status
        else:
            st = all_status.get(key)
        cm = all_comment.get(key)
        has_img = any(
            allowed_file(f) and f.startswith(f"test_{group}_{key}_")
            for f in file_list
        )
        test_status[key] = {
            'status': st,
            'comment': cm,
            'has_img': has_img
        }


    return render_template(
        "test_group_menu.html",
        report=report,
        group=group,
        test_titles=group_titles,
        test_status=test_status
    )

@app.route('/test_group/<report>/<group>/<test_key>', methods=['GET', 'POST'])
def test_group_item_dynamic(report, group, test_key):
    session[f"last_test_type_{report}"] = get_group_title(group)
    # Nếu là hot_cold cycle test thì redirect sang route mới
    if test_key == "hot_cold" and group in ["indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal"]:
        return redirect(url_for("hot_cold_test", report=report, group=group))

    # Lấy cấu hình group
    group_titles = TEST_GROUP_TITLES.get(group)
    if not group_titles or test_key not in group_titles:
        return "Mục kiểm tra không tồn tại!", 404
    title = group_titles[test_key]

    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(report_folder, exist_ok=True)
    status_file = os.path.join(report_folder, f"status_{group}.txt")
    comment_file = os.path.join(report_folder, f"comment_{group}.txt")

    from test_logic import update_group_note_file, get_group_note_value

    # --- Trạng thái PASS/FAIL/N/A ---
    all_status = load_group_notes(status_file)
    status_value = all_status.get(test_key, "")

    # --- Comment ---
    def update_comment(file_path, key, value):
        update_group_note_file(file_path, key, value)

    def get_comment(file_path, key):
        return get_group_note_value(file_path, key)

    # Lấy comment của mục này
    comment = get_comment(comment_file, test_key)

    # --- Xác định loại test đặc biệt ---
    is_rh_np = (group == "transit_RH_np")
    is_drop = is_drop_test(title) if group.startswith("transit") else False
    is_impact = is_impact_test(title) if group.startswith("transit") else False
    is_rot = is_rotational_test(title) if group.startswith("transit") else False

    # --- RH Non Pallet zones ---
    rh_impact_zones = RH_IMPACT_ZONES if is_rh_np and test_key == "step3" else []
    rh_vib_zones = RH_VIB_ZONES if is_rh_np and test_key == "step4" else []
    rh_second_impact_zones = RH_SECOND_IMPACT_ZONES if is_rh_np and test_key == "step5" else []
    rh_step12_zones = RH_STEP12_ZONES if is_rh_np and test_key == "step12" else []

    # --- Xử lý upload ảnh, xóa ảnh, comment, status ---
    if request.method == 'POST':
        # Upload vùng RH (step3/4/5)
        for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
            files = request.files.getlist(f'rh_impact_img_{zone}') or \
                    request.files.getlist(f'rh_vib_img_{zone}') or \
                    request.files.getlist(f'rh_step12_img_{zone}') or \
                    request.files.getlist(f'rh_second_impact_img_{zone}')
            for file in files:
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[-1].lower()
                    prefix = f"test_{group}_{test_key}_{zone}_"
                    current_nums = [
                        int(f[len(prefix):].split('.')[0])
                        for f in os.listdir(report_folder)
                        if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()
                    ]
                    next_num = max(current_nums) + 1 if current_nums else 1
                    new_fname = f"{prefix}{next_num}.{ext}"
                    file.save(os.path.join(report_folder, new_fname))
        # Xử lý upload ảnh thường
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
        # Xóa ảnh thường hoặc vùng
        if 'delete_img' in request.form:
            del_img = request.form['delete_img']
            img_path = os.path.join(report_folder, del_img)
            if os.path.exists(img_path):
                os.remove(img_path)
        # Ghi status PASS/FAIL/N/A
        if 'status' in request.form:
            update_group_note_file(status_file, test_key, request.form['status'])
            status = request.form['status']

        # Ghi comment
        if 'save_comment' in request.form:
            comment_val = request.form.get('comment_input', '').strip()
            update_comment(comment_file, test_key, comment_val)
            comment = comment_val

        return redirect(request.url)

    # --- Chuẩn bị dữ liệu ảnh vùng RH (step3/4/5) ---
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
        if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_") and all(not f.startswith(f"test_{group}_{test_key}_{zone}_") for zone, _ in rh_impact_zones + rh_vib_zones + rh_second_impact_zones):
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

    # Ưu tiên RH non pallet step3, step4, step5, các bước còn lại mới nhận diện như cũ!
    is_rh_np = (group == "transit_RH_np")
    is_rh_np_step3 = is_rh_np and key == "step3"
    is_rh_np_step4 = is_rh_np and key == "step4"
    is_rh_np_step5 = is_rh_np and key == "step5"

    # Chỉ step3 RH non pallet mới có 9 vùng impact riêng
    rh_impact_zones = RH_IMPACT_ZONES if is_rh_np_step3 else []
    rh_vib_zones = RH_VIB_ZONES if is_rh_np_step4 else []
    rh_second_impact_zones = RH_SECOND_IMPACT_ZONES if is_rh_np_step5 else []
    allow_na = is_rh_np and (key in ["step6", "step7", "step8", "step11", "step12"])

    # === Xử lý ảnh từng vùng RH non pallet (nếu là step3, step4, step5) ===
    zone_imgs = {}
    for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
        imgs = []
        if os.path.exists(report_folder):
            for f in os.listdir(report_folder):
                # Đặt tên file: test_{group}_{key}_{zone}_...
                if allowed_file(f) and f.startswith(f"test_{group}_{key}_{zone}_"):
                    imgs.append(f"/images/{report}/{f}")
        zone_imgs[zone] = imgs

    # === Xử lý nhận diện các loại test còn lại (impact, drop, rot) ===
    # Định nghĩa nhóm transit 2C
    TRANSIT_2C_GROUPS = ("transit_2c_np", "transit_2c_pallet")

    if not (is_rh_np_step3 or is_rh_np_step4 or is_rh_np_step5):
        is_transit_2c = group in TRANSIT_2C_GROUPS
        is_drop = is_drop_test(title) and is_transit_2c
        is_impact = is_impact_test(title) and is_transit_2c
        is_rot = is_rotational_test(title) and is_transit_2c
    else:
        is_drop = is_impact = is_rot = False

    # === Xử lý upload ảnh cho các loại test thông thường ===
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

    # === Trạng thái từng mục (PASS/FAIL/NA, comment,...) ===
    def update_group_note_file(file_path, key, value):
        lines = []
        found = False
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        new_lines = []
        for line in lines:
            if line.strip().startswith(f"Mục {key}:"):
                new_lines.append(f"Mục {key}: {value}\n")
                found = True
            else:
                new_lines.append(line)
        if not found:
            new_lines.append(f"Mục {key}: {value}\n")
        with open(file_path, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)

    def get_group_note_value(file_path, key):
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.strip().startswith(f"Mục {key}:"):
                        return line.strip().split(":", 1)[1].strip()
        return None

    status_value = get_group_note_value(status_file, key)

    # === Xử lý POST (upload ảnh, xóa, comment, status) ===
    if request.method == 'POST':
        # Ưu tiên upload các vùng RH non pallet step3/4/5 trước
        # Xử lý upload ảnh cho từng vùng
        for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
            files = request.files.getlist(f'rh_impact_img_{zone}')
            for file in files:
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[-1].lower()
                    prefix = f"test_{group}_{key}_{zone}_"
                    current_nums = [
                        int(f[len(prefix):].split('.')[0])
                        for f in os.listdir(report_folder)
                        if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()
                    ]
                    next_num = max(current_nums) + 1 if current_nums else 1
                    new_fname = f"{prefix}{next_num}.{ext}"
                    file.save(os.path.join(report_folder, new_fname))
        # Xử lý xóa ảnh từng vùng
        if 'delete_img' in request.form:
            del_img = request.form['delete_img']
            img_path = os.path.join(report_folder, del_img)
            if os.path.exists(img_path):
                os.remove(img_path)
        # Status và comment
        if 'status' in request.form and not group.startswith("transit"):
            status = request.form['status']
            update_group_note_file(status_file, key, status)
        if 'save_comment' in request.form:
            comment = request.form.get('comment_input', '').strip()
            update_group_note_file(comment_file, key, comment)
        # Upload ảnh dạng thường nếu không phải drop
        if 'test_imgs' in request.files and not is_drop:
            files = request.files.getlist('test_imgs')
            for file in files:
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[-1].lower()
                    prefix = f"test_{group}_{key}_"
                    current_nums = [
                        int(f[len(prefix):].split('.')[0])
                        for f in os.listdir(report_folder)
                        if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()
                    ]
                    next_num = max(current_nums) + 1 if current_nums else 1
                    new_fname = f"test_{group}_{key}_{next_num}.{ext}"
                    file.save(os.path.join(report_folder, new_fname))
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
        # -- Thêm logic chọn template theo từng bước --
    TRANSIT_GROUPS = ("transit_2c_np", "transit_2c_pallet", "transit_RH_np", "transit_RH_pallet")
    transit_special_steps = []
    if group == "transit_RH_np":
        transit_special_steps = ["step3", "step4", "step5", "step12"]
    elif group == "transit_2c_np":
        transit_special_steps = ["step6"]  # (bổ sung step đặc biệt khác nếu có)
    elif group == "transit_2c_pallet":
        transit_special_steps = ["step8", "step9"]  # (ví dụ, bổ sung các step cần vùng đặc biệt)
    elif group == "transit_RH_pallet":
        transit_special_steps = ["step3", "step6"]  # (bổ sung nếu cần)
    # ... (mỗi group tùy nhu cầu bổ sung step)

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
        # ---- RH non pallet đặc biệt:
        is_rh_np=is_rh_np,
        rh_impact_zones=rh_impact_zones,
        rh_vib_zones=rh_vib_zones,
        rh_second_impact_zones=rh_second_impact_zones,
        allow_na=allow_na,
        zone_imgs=zone_imgs,
    )

@app.route("/hot_cold_test/<report>/<group>", methods=["GET", "POST"])
def hot_cold_test(report, group):
    session[f"last_test_type_{report}"] = f"HOT & COLD CYCLE TEST ({group.upper()})"
    vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(folder, exist_ok=True)
    status_file = os.path.join(folder, f"hotcold_status_{group}.txt")
    comment_file = os.path.join(folder, f"hotcold_comment_{group}.txt")
    before_tag = f"hotcold_before_{group}"
    after_tag = f"hotcold_after_{group}"
    before_time_file = os.path.join(folder, f"hotcold_before_time_{group}.txt")
    after_time_file = os.path.join(folder, f"hotcold_after_time_{group}.txt")

    # --- Xử lý POST ---
    if request.method == "POST":
        if "status" in request.form:
            with open(status_file, "w", encoding="utf-8") as f:
                f.write(request.form["status"])
        if "save_comment" in request.form:
            with open(comment_file, "w", encoding="utf-8") as f:
                f.write(request.form.get("comment_input", ""))
        for tag, time_file in [(before_tag, before_time_file), (after_tag, after_time_file)]:
            if f"{tag}_imgs" in request.files:
                files = request.files.getlist(f"{tag}_imgs")
                count = 0
                for file in files:
                    if file and allowed_file(file.filename):
                        ext = file.filename.rsplit('.', 1)[-1].lower()
                        prefix = f"{tag}_"
                        nums = [
                            int(f[len(prefix):].split('.')[0])
                            for f in os.listdir(folder)
                            if f.startswith(prefix) and f[len(prefix):].split('.')[0].isdigit()
                        ]
                        next_num = max(nums) + 1 if nums else 1
                        new_fname = f"{tag}_{next_num}.{ext}"
                        file.save(os.path.join(folder, new_fname))
                        count += 1
                if count > 0:
                    now = datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M")
                    with open(time_file, "w", encoding="utf-8") as tf:
                        tf.write(now)
        if "delete_img" in request.form:
            img = request.form["delete_img"]
            img_path = os.path.join(folder, img)
            if os.path.exists(img_path):
                os.remove(img_path)
            # Xử lý xóa file time nếu không còn ảnh
            if img.startswith(before_tag):
                still_imgs = [f for f in os.listdir(folder) if allowed_file(f) and f.startswith(before_tag)]
                if not still_imgs and os.path.exists(before_time_file):
                    os.remove(before_time_file)
            if img.startswith(after_tag):
                still_imgs = [f for f in os.listdir(folder) if allowed_file(f) and f.startswith(after_tag)]
                if not still_imgs and os.path.exists(after_time_file):
                    os.remove(after_time_file)
        session[f"last_test_type_{report}"] = f"HOT & COLD CYCLE TEST ({group.upper()})"
        return redirect(request.url)

    # --- Lấy trạng thái và comment ---
    status = ""
    comment = ""
    if os.path.exists(status_file):
        with open(status_file, "r", encoding="utf-8") as f:
            status = f.read().strip()
    if os.path.exists(comment_file):
        with open(comment_file, "r", encoding="utf-8") as f:
            comment = f.read().strip()
    test_key = "hot_cold"
    imgs_mo_ta = TEST_GROUP_TITLES[group][test_key]["img"]
    # --- Lấy danh sách ảnh before/after ---
    imgs_before = []
    imgs_after = []
    for f in sorted(os.listdir(folder)):
        if allowed_file(f):
            if f.startswith(before_tag):
                imgs_before.append(f"/images/{report}/{f}")
            if f.startswith(after_tag):
                imgs_after.append(f"/images/{report}/{f}")
    # --- Lấy thời gian upload nếu có ---
    before_upload_time = None
    after_upload_time = None
    if os.path.exists(before_time_file):
        with open(before_time_file, "r", encoding="utf-8") as f:
            before_upload_time = f.read().strip()
    if os.path.exists(after_time_file):
        with open(after_time_file, "r", encoding="utf-8") as f:
            after_upload_time = f.read().strip()

    return render_template(
        "hot_cold_test.html",  # Template riêng cho hot-cold, hoặc dùng lại LINE_TEST_TEMPLATE đều được!
        report=report,
        group=group,
        status=status,
        comment=comment,
        imgs_mo_ta=imgs_mo_ta,
        imgs_before=imgs_before,
        imgs_after=imgs_after,
        before_upload_time=before_upload_time,
        after_upload_time=after_upload_time,
        so_gio_test=SO_GIO_TEST,
    )

def get_hotcold_elapsed(report, group):
    # Giả sử bạn lưu time file ở: images/report/hot_cold_upload_time.txt
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    time_file = os.path.join(folder, f"hot_cold_upload_time_{group}.txt")
    if os.path.exists(time_file):
        with open(time_file, "r", encoding="utf-8") as f:
            tstr = f.read().strip()
        try:
            vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
            dt = datetime.strptime(tstr, "%d/%m/%Y %H:%M")
            dt = vn_tz.localize(dt)
            now = datetime.now(vn_tz)
            return (now - dt).total_seconds() / 3600
        except:
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

    def read_textfile(path):
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return f.read().strip()
        return ""
    def write_textfile(path, text):
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)

    # --- POST ---
    if request.method == "POST":
        # Lưu trạng thái PASS/FAIL/DATA
        if "status" in request.form:
            write_textfile(files_map["status"], request.form["status"])
            if request.form["status"] != "FAIL":
                if os.path.exists(files_map["comment"]): os.remove(files_map["comment"])
        # Lưu fail reason
        if "save_fail_reason" in request.form:
            reasons = request.form.getlist("fail_reason")
            other = request.form.get("fail_reason_other", "").strip()
            if other: reasons.append(other)
            write_textfile(files_map["comment"], "; ".join(reasons))
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
                    write_textfile(time_file, datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M"))
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
    status = read_textfile(files_map["status"])
    fail_reason_raw = read_textfile(files_map["comment"])
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
    before_upload_time = read_textfile(files_map["before_time"])
    after_upload_time  = read_textfile(files_map["after_time"])

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
    vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    before_time_file = os.path.join(folder, "before_upload_time.txt")
    if os.path.exists(before_time_file):
        with open(before_time_file, "r", encoding="utf-8") as f:
            tstr = f.read().strip()
        try:
            dt = datetime.strptime(tstr, "%d/%m/%Y %H:%M")
            now = datetime.now(vn_tz)
            elapsed = (now - vn_tz.localize(dt)).total_seconds() / 3600
            return elapsed
        except Exception as e:
            print("Parse time error:", e)
            return None
    return None

@app.route("/store_sample", methods=["GET", "POST"])
def store_sample():
    report = request.args.get("report")
    item_code = get_item_code(report)
    auto_sample_name = f"{report} - {item_code}" if report and item_code else ""
    error_msg = ""

    # === Kiểm tra đã có mẫu lưu với report+item_code này chưa ===
    found_location = None
    for loc, info in SAMPLE_STORAGE.items():
        if info.get("report") == report and info.get("item_code") == item_code:
            found_location = loc
            break

    if found_location:
        # Đã có mẫu => chuyển sang trang infor của mẫu đó (hoặc render thông tin + message)
        return redirect(url_for("sample_map", location_id=found_location))

    # === Nếu chưa có thì xử lý như cũ ===
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
    if request.method == "POST":
        loc = request.form.get("loc")
        borrower = request.form.get("borrower")
        note = request.form.get("note")
        if loc in SAMPLE_STORAGE:
            SAMPLE_STORAGE[loc]['borrower'] = borrower
            SAMPLE_STORAGE[loc]['note'] = note

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
    return send_from_directory(folder, filename)

@app.route("/view_counter_log")
def view_counter_log():
    import openpyxl
    from collections import defaultdict, OrderedDict

    excel_path = "counter_detail_log.xlsx"
    rows = []
    type_of_set = set()
    ca_map = {"office": "HC", "hc": "HC", "ot": "OT"}
    # Đọc file
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        col_idx = {str(cell.value).strip().lower(): i for i, cell in enumerate(ws[1], 0)}
        date_idx = col_idx.get("ngày", 0)
        ca_idx = col_idx.get("ca", 2)
        type_idx = col_idx.get("type of", 4)

        # Gom dữ liệu thành: {ngày: {ca: {type_of_short: số lượng}}}
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

        # Chỉ giữ 10 ngày mới nhất
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
    else:
        header = ["Ngày", "Ca", "Tổng"]

    return render_template(
        "counter_log.html",
        header=header,
        rows=rows,
        type_of_list=type_of_list if rows else [],
    )

def auto_notify_all_first_time():
    webhook_url = TEAMS_WEBHOOK_URL
    for report_folder in os.listdir(UPLOAD_FOLDER):
        folder = os.path.join(UPLOAD_FOLDER, report_folder)
        if not os.path.isdir(folder): continue

        # Line test: gửi ngay khi đủ giờ (job mỗi phút)
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
        # Hotcold test: gửi ngay khi đủ giờ (job mỗi phút)
        for group in ["indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal"]:
            notify_when_enough_time(
                report=report_folder,
                so_gio_test=SO_GIO_TEST,
                tag_after="hot_cold_after",
                time_file_name=f"hotcold_before_time_{group}.txt",
                flag_file_name=f"teams_notified_hotcold_{group}.txt",
                webhook_url=webhook_url,
                notify_msg=f"✅ [TỰ ĐỘNG] Hot & Cold cycle test của REPORT {report_folder} ({group.upper()}) đã đủ {SO_GIO_TEST} tiếng! Vui lòng upload ảnh after.",
                force_send=False,
                pending_notify_name=f"pending_notify_hotcold_{group}.txt"
            )

def auto_notify_all_repeat():
    webhook_url = TEAMS_WEBHOOK_URL
    for report_folder in os.listdir(UPLOAD_FOLDER):
        folder = os.path.join(UPLOAD_FOLDER, report_folder)
        if not os.path.isdir(folder): continue

        # Line test: lặp lại mỗi 1 tiếng nếu chưa có ảnh after
        notify_when_enough_time(
            report=report_folder,
            so_gio_test=SO_GIO_TEST,
            tag_after="line_after",
            time_file_name="before_upload_time.txt",
            flag_file_name=None,   # Không dùng flag ở lần lặp lại
            webhook_url=webhook_url,
            notify_msg=f"✅ [TỰ ĐỘNG, NHẮC LẠI 1 TIẾNG] Line test của sản phẩm REPORT {report_folder} đã đủ {SO_GIO_TEST} tiếng! Vui lòng upload ảnh after.",
            force_send=True,
            pending_notify_name="pending_notify_line.txt"
        )
        # Hotcold test: lặp lại mỗi 1 tiếng nếu chưa có ảnh after
        for group in ["indoor_chuyen", "indoor_thuong", "indoor_stone", "indoor_metal"]:
            notify_when_enough_time(
                report=report_folder,
                so_gio_test=SO_GIO_TEST,
                tag_after="hot_cold_after",
                time_file_name=f"hotcold_before_time_{group}.txt",
                flag_file_name=None,
                webhook_url=webhook_url,
                notify_msg=f"✅ [TỰ ĐỘNG, NHẮC LẠI 1 TIẾNG] Hot & Cold cycle test của REPORT {report_folder} ({group.upper()}) đã đủ {SO_GIO_TEST} tiếng! Vui lòng upload ảnh after.",
                force_send=True,
                pending_notify_name=f"pending_notify_hotcold_{group}.txt"
            )

def auto_notify_all_pending():
    webhook_url = TEAMS_WEBHOOK_URL
    now = datetime.now()
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
            pending_path = os.path.join(folder, f"pending_notify_hotcold_{group}.txt")
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

if __name__ == "__main__":
    app.run("0.0.0.0", port=8080, debug=True)
