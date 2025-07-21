# app.py
from flask import Flask, request, render_template, session, redirect, url_for
from config import SECRET_KEY, PASSWORD, local_main, SAMPLE_STORAGE, UPLOAD_FOLDER, TEST_GROUPS, local_complete, qr_folder, SO_GIO_TEST, ALL_SLOTS
from excel_utils import get_item_code, get_col_idx, copy_row_with_style, is_img_at_cell
from image_utils import allowed_file, safe_filename, get_img_urls
from auth import login, logout, is_logged_in
from test_logic import load_group_notes, get_group_test_status, is_drop_test, is_impact_test, is_rotational_test,  TEST_GROUP_TITLES, TEST_TYPE_VI, DROP_ZONES, DROP_LABELS
from test_logic import IMPACT_ZONES, IMPACT_LABELS, ROT_LABELS, ROT_ZONES, RH_IMPACT_ZONES, RH_VIB_ZONES, RH_SECOND_IMPACT_ZONES, update_group_note_file, get_group_note_value
from notify_utils import send_teams_message
from counter_utils import update_counter, log_order_complete, check_and_reset_counter, log_report_complete
from openpyxl import load_workbook, Workbook
from flask import send_from_directory
from datetime import datetime
from openpyxl.styles import PatternFill
from excel_utils import ensure_column
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import re
import os
import pytz
import json
import openpyxl
from collections import defaultdict, OrderedDict

app = Flask(__name__)
app.secret_key = SECRET_KEY

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
            wb = load_workbook(local_main)
            ws = wb.active
            rating_col = get_col_idx(ws, "rating")
            status_col = get_col_idx(ws, "status")
            ws.cell(row=row_idx, column=rating_col).value = value

            if status_col:
                ws.cell(row=row_idx, column=status_col).value = "COMPLETE"
                fill_complete = PatternFill("solid", fgColor="BFBFBF")
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col).fill = fill_complete

            # Copy sang completed file
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

    # === Gửi thông báo Teams khi đủ giờ và chưa gửi ===
    webhook_url = "https://mphcmiuedu.webhook.office.com/webhookb2/49e44c2c-a806-4877-8cc7-951b25b18a86@a7380202-eb54-415a-9b66-4d9806cfab42/IncomingWebhook/e297177041eb426bbbd78da7c07da4e8/81e98e9d-7a3f-492d-9f99-3ac5a5ecbbf3/V2JMRXicdFXVHV5Ouh8uLnNdLE11NvOje_PfEwGNG_zoM1"
    notified_flag = os.path.join(UPLOAD_FOLDER, str(report), "teams_notified.txt")
    if show_line_test_done and not os.path.exists(notified_flag):
        ok = send_teams_message(
            webhook_url,
            f"✅ [Thông báo tự động] Line test của sản phẩm REPORT {report} đã hoàn thành {SO_GIO_TEST} tiếng!"
        )
        if ok:
            with open(notified_flag, "w") as f:
                f.write("1")
        else:
            print("Lỗi gửi Teams: Không gửi được thông báo!")

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
        show_line_test_done_notice=show_line_test_done_notice,
        so_gio_test=SO_GIO_TEST,
    )

@app.route("/test_group/<report>/<group>", methods=["GET", "POST"])
def test_group_page(report, group):
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
        test_status[key] = {
            'status': all_status.get(key),
            'comment': all_comment.get(key),
            'has_img': any(
                allowed_file(f) and f.startswith(f"test_{group}_{key}_")
                for f in file_list
            )
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

    # --- Trạng thái PASS/FAIL/N/A ---
    all_status = load_group_notes(status_file)
    status_value = all_status.get(test_key, "")

    # --- Comment ---
    def update_comment(file_path, value):
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(value)
    def get_comment(file_path):
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read().strip()
        return ""
    comment = get_comment(comment_file)

    # --- Xác định loại test đặc biệt ---
    is_rh_np = (group == "transit_RH_np")
    is_drop = is_drop_test(title) if group.startswith("transit") else False
    is_impact = is_impact_test(title) if group.startswith("transit") else False
    is_rot = is_rotational_test(title) if group.startswith("transit") else False

    # --- RH Non Pallet zones ---
    rh_impact_zones = RH_IMPACT_ZONES if is_rh_np and test_key == "step3" else []
    rh_vib_zones = RH_VIB_ZONES if is_rh_np and test_key == "step4" else []
    rh_second_impact_zones = RH_SECOND_IMPACT_ZONES if is_rh_np and test_key == "step5" else []

    # --- Xử lý upload ảnh, xóa ảnh, comment, status ---
    if request.method == 'POST':
        # Upload vùng RH (step3/4/5)
        for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
            files = request.files.getlist(f'rh_impact_img_{zone}') or \
                    request.files.getlist(f'rh_vib_img_{zone}') or \
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
        # Ghi comment
        if 'save_comment' in request.form:
            comment = request.form.get('comment_input', '').strip()
            update_comment(comment_file, comment)
        return redirect(request.url)

    # --- Chuẩn bị dữ liệu ảnh vùng RH (step3/4/5) ---
    zone_imgs = {}
    for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
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
    return render_template(
        "test_transit_item.html",
        report=report,
        group=group,
        key=test_key,
        title=title,
        is_rh_np=is_rh_np,
        status=status_value,
        imgs=imgs,
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
        rh_impact_zones=rh_impact_zones,
        rh_vib_zones=rh_vib_zones,
        rh_second_impact_zones=rh_second_impact_zones,
        zone_imgs=zone_imgs,
        comment=comment,
    )

def render_test_group_item(report, group, key, group_titles):
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
    # CHỈ NHẬN DIỆN khi KHÔNG phải RH non pallet step3/step4/step5!
    if not (is_rh_np_step3 or is_rh_np_step4 or is_rh_np_step5):
        is_drop = is_drop_test(title) if group.startswith("transit") else False
        is_impact = is_impact_test(title) if group.startswith("transit") else False
        is_rot = is_rotational_test(title) if group.startswith("transit") else False
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
    template_name = "test_transit_item.html" if group.startswith("transit") else "test_group_item.html"

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
    
@app.route("/line_test/<report>", methods=["GET", "POST"])
def line_test(report):
    folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(folder, exist_ok=True)
    status_file = os.path.join(folder, "line_status.txt")
    comment_file = os.path.join(folder, "line_comment.txt")
    before_tag = "line_before"
    after_tag = "line_after"
    before_time_file = os.path.join(folder, "before_upload_time.txt")
    after_time_file = os.path.join(folder, "after_upload_time.txt")

    # Danh sách nguyên nhân fail mẫu
    fail_reasons_list = [
        "Vật liệu bị ẩm.",
        "Vị trí bị tách lớp, mặt dưới veneer có phủ keo.",
        "Vị trí bị tách lớp, mặt dưới veneer không phủ đều keo."
    ]

    # Xử lý POST
    if request.method == "POST":
        if "status" in request.form:
            with open(status_file, "w", encoding="utf-8") as f:
                f.write(request.form["status"])
            # Nếu chọn PASS hoặc DATA thì xóa nguyên nhân fail cũ (nếu có)
            if request.form["status"] != "FAIL" and os.path.exists(comment_file):
                os.remove(comment_file)
        # Nếu là lưu nguyên nhân fail (nhấn nút lưu ở form fail reason)
        if "save_fail_reason" in request.form:
            reasons = request.form.getlist("fail_reason")
            other = request.form.get("fail_reason_other", "").strip()
            reasons_full = reasons.copy()
            if other:
                reasons_full.append(other)
            reasons_text = "; ".join(reasons_full)
            with open(comment_file, "w", encoding="utf-8") as f:
                f.write(reasons_text)
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
                # Nếu có upload ảnh mới => cập nhật lại thời gian (ghi đè luôn)
                if count > 0:
                    with open(time_file, "w", encoding="utf-8") as tf:
                        vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
                        now = datetime.now(vn_tz).strftime("%d/%m/%Y %H:%M")
                        tf.write(now)
        if "delete_img" in request.form:
            img = request.form["delete_img"]
            img_path = os.path.join(folder, img)
            if os.path.exists(img_path):
                os.remove(img_path)
            # Nếu là ảnh before => check còn ảnh không, nếu không còn thì xóa file time
            if img.startswith(before_tag):
                still_imgs = [f for f in os.listdir(folder) if allowed_file(f) and f.startswith(before_tag)]
                if not still_imgs and os.path.exists(before_time_file):
                    os.remove(before_time_file)
            # Nếu là ảnh after => check còn ảnh không, nếu không còn thì xóa file time
            if img.startswith(after_tag):
                still_imgs = [f for f in os.listdir(folder) if allowed_file(f) and f.startswith(after_tag)]
                if not still_imgs and os.path.exists(after_time_file):
                    os.remove(after_time_file)
        session[f"last_test_type_{report}"] = "LINE TEST"
        return redirect(request.url)

    # Lấy trạng thái và nguyên nhân fail đã lưu
    status = ""
    fail_reasons = []
    fail_reason_other = ""
    if os.path.exists(status_file):
        with open(status_file, "r", encoding="utf-8") as f:
            status = f.read().strip()
    if os.path.exists(comment_file):
        with open(comment_file, "r", encoding="utf-8") as f:
            text = f.read().strip()
            # Phân biệt nguyên nhân khác (không nằm trong fail_reasons_list)
            reasons_all = [s.strip() for s in text.split(";") if s.strip()]
            for r in reasons_all.copy():
                if r not in fail_reasons_list:
                    fail_reason_other = r
                    reasons_all.remove(r)
            fail_reasons = reasons_all

    # Lấy danh sách ảnh before/after
    imgs_before = []
    imgs_after = []
    for f in sorted(os.listdir(folder)):
        if allowed_file(f):
            if f.startswith(before_tag):
                imgs_before.append(f"/images/{report}/{f}")
            if f.startswith(after_tag):
                imgs_after.append(f"/images/{report}/{f}")

    # Lấy thời gian upload nếu có
    before_upload_time = None
    after_upload_time = None
    if os.path.exists(before_time_file):
        with open(before_time_file, "r", encoding="utf-8") as f:
            before_upload_time = f.read().strip()
    if os.path.exists(after_time_file):
        with open(after_time_file, "r", encoding="utf-8") as f:
            after_upload_time = f.read().strip()

    return render_template(
        "line_test.html",
        report=report,
        status=status,
        fail_reasons=fail_reasons,            # Tick sẵn các nguyên nhân mẫu
        fail_reason_other=fail_reason_other,  # Dòng nhập "khác"
        fail_reasons_list=fail_reasons_list,  # Danh sách tick mẫu
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

@app.route('/test_group/<report>/<group>/<step_key>', methods=['GET', 'POST'])
def transit_item_page(report, group, step_key):
    """
    Giao diện test từng bước (item) transit, có vùng upload ảnh, comment, PASS/FAIL/N/A và các vùng đặc biệt (drop/impact/RH...).
    """
    from test_logic import (
        is_drop_test, is_impact_test, is_rotational_test, load_group_notes, 
        RH_IMPACT_ZONES, RH_VIB_ZONES, RH_SECOND_IMPACT_ZONES, 
        DROP_LABELS, DROP_ZONES, IMPACT_LABELS, IMPACT_ZONES,
        ROT_LABELS, ROT_ZONES, TEST_GROUP_TITLES
    )
    group_titles = TEST_GROUP_TITLES.get(group)
    if not group_titles or step_key not in group_titles:
        return "Không tìm thấy bước kiểm tra!", 404
    title = group_titles[step_key]

    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    os.makedirs(report_folder, exist_ok=True)
    status_file = os.path.join(report_folder, f"status_{group}.txt")
    comment_file = os.path.join(report_folder, f"comment_{group}.txt")

    # --- Trạng thái PASS/FAIL/N/A ---
    all_status = load_group_notes(status_file)
    status_value = all_status.get(step_key, "")

    # --- Comment ---
    def update_comment(file_path, value):
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(value)
    def get_comment(file_path):
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read().strip()
        return ""
    comment = get_comment(comment_file)

    # --- Xác định loại test đặc biệt ---
    is_rh_np = (group == "transit_RH_np")
    is_drop = is_drop_test(title) if group.startswith("transit") else False
    is_impact = is_impact_test(title) if group.startswith("transit") else False
    is_rot = is_rotational_test(title) if group.startswith("transit") else False

    # --- RH Non Pallet zones ---
    rh_impact_zones = RH_IMPACT_ZONES if is_rh_np and step_key == "step3" else []
    rh_vib_zones = RH_VIB_ZONES if is_rh_np and step_key == "step4" else []
    rh_second_impact_zones = RH_SECOND_IMPACT_ZONES if is_rh_np and step_key == "step5" else []

    # --- Xử lý upload ảnh, xóa ảnh, comment, status ---
    if request.method == 'POST':
        # Upload vùng RH (step3/4/5)
        for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
            files = request.files.getlist(f'rh_impact_img_{zone}') or \
                    request.files.getlist(f'rh_vib_img_{zone}') or \
                    request.files.getlist(f'rh_second_impact_img_{zone}')
            for file in files:
                if file and allowed_file(file.filename):
                    ext = file.filename.rsplit('.', 1)[-1].lower()
                    prefix = f"test_{group}_{step_key}_{zone}_"
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
                    prefix = f"test_{group}_{step_key}_"
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
            from test_logic import update_group_note_file
            update_group_note_file(status_file, step_key, request.form['status'])
        # Ghi comment
        if 'save_comment' in request.form:
            comment = request.form.get('comment_input', '').strip()
            update_comment(comment_file, comment)
        return redirect(request.url)

    # --- Chuẩn bị dữ liệu ảnh vùng RH (step3/4/5) ---
    zone_imgs = {}
    for zone, label in rh_impact_zones + rh_vib_zones + rh_second_impact_zones:
        imgs_zone = []
        for f in os.listdir(report_folder):
            if allowed_file(f) and f.startswith(f"test_{group}_{step_key}_{zone}_"):
                imgs_zone.append(f"/images/{report}/{f}")
        zone_imgs[zone] = imgs_zone

    # --- Chuẩn bị dữ liệu ảnh thường ---
    imgs = []
    for f in sorted(os.listdir(report_folder)):
        if allowed_file(f) and f.startswith(f"test_{group}_{step_key}_") and all(not f.startswith(f"test_{group}_{step_key}_{zone}_") for zone, _ in rh_impact_zones + rh_vib_zones + rh_second_impact_zones):
            imgs.append(f"/images/{report}/{f}")

    # --- Chuẩn bị ảnh drop, impact, rot nếu có ---
    drop_imgs, impact_imgs, rot_imgs = [], [], []
    if is_drop:
        for zone in DROP_ZONES:
            di = []
            for f in os.listdir(report_folder):
                if allowed_file(f) and f.startswith(f"test_{group}_{step_key}_drop_{zone}_"):
                    di.append(f"/images/{report}/{f}")
            drop_imgs.append(di)
    if is_impact:
        for zone in IMPACT_ZONES:
            ii = []
            for f in os.listdir(report_folder):
                if allowed_file(f) and f.startswith(f"test_{group}_{step_key}_impact_{zone}_"):
                    ii.append(f"/images/{report}/{f}")
            impact_imgs.append(ii)
    if is_rot:
        for zone in ROT_ZONES:
            ri = []
            for f in os.listdir(report_folder):
                if allowed_file(f) and f.startswith(f"test_{group}_{step_key}_rotation_{zone}_"):
                    ri.append(f"/images/{report}/{f}")
            rot_imgs.append(ri)

    # --- Trả về template ---
    return render_template(
        "test_transit_item.html",
        report=report,
        group=group,
        key=step_key,
        title=title,
        is_rh_np=is_rh_np,
        status=status_value,
        imgs=imgs,
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
        rh_impact_zones=rh_impact_zones,
        rh_vib_zones=rh_vib_zones,
        rh_second_impact_zones=rh_second_impact_zones,
        zone_imgs=zone_imgs,
        comment=comment,
    )

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

if __name__ == "__main__":
    app.run("0.0.0.0", port=8080, debug=True)
