# docx_utils.py
import os
import re
import time
import uuid
import tempfile
from openpyxl import load_workbook
from docx import Document
from config import local_main

# === CONFIG ===
WORD_TEMPLATE = "FORM-QAD-011-TEST REQUEST FORM (TRF).docx"
PDF_OUTPUT_FOLDER = os.path.join("static", "TFR")

# =========================
# Blank-detector for Excel
# =========================
BLANK_TOKENS = {"", "-", "—"}

def _normalize_to_check_blank(v):
    if v is None:
        return True, ""
    if isinstance(v, str):
        s = (
            v.replace("\u00A0", "")
             .replace("\u200B", "")
             .replace("\r", "")
             .replace("\n", "")
             .replace("\t", "")
             .strip()
        )
        return (s in BLANK_TOKENS), s
    return False, str(v)

def get_first_empty_report_all_blank(excel_path):
    """Tìm mã report đầu tiên có toàn bộ cột C..X đều trống."""
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active
    for row in range(2, ws.max_row + 1):
        all_mid_empty = True
        for col in range(3, 25):  # C..X
            is_blank, _ = _normalize_to_check_blank(ws.cell(row=row, column=col).value)
            if not is_blank:
                all_mid_empty = False
                break
        if all_mid_empty:
            report_no = ws.cell(row=row, column=2).value  # cột B
            if report_no is not None and str(report_no).strip() != "":
                wb.close()
                return str(report_no).strip()
    wb.close()
    return None

# =========================
# Checkbox mapping cho DOCX
# =========================
def build_label_value_map(data):
    """
    Sinh map {label_in_template: bool} cho các checkbox trong file Word.
    - Nới lỏng 'test_group': khớp cả “CONSTRUCTION” và “CONSTRUCTION TEST”.
    - Giữ logic '...TH' cho test_status (nth).
    - Sinh cờ N/A cho các trường văn bản.
    """
    label_groups = {
        "test_group": [
            "MATERIAL TEST",
            "FINISHING TEST",
            "CONSTRUCTION TEST",
            "TRANSIT TEST",
            "ENVIRONMENTAL TEST",
        ],
        "test_status": ["1ST", "2ND", "3RD", "...TH"],  # nth → ...TH
        "furniture_testing": ["INDOOR", "OUTDOOR"],
        "sample_return": ["YES", "NO"],
    }

    def _eq_relaxed(label: str, value: str, group: str) -> bool:
        L = (label or "").strip().upper()
        V = (value or "").strip().upper()
        if not V:
            return False
        if group == "test_group":
            # chấp nhận có/không “ TEST” ở value hoặc label
            if V.endswith(" TEST"):
                V2 = V[:-5].strip()
            else:
                V2 = V
            if L.endswith(" TEST"):
                L2 = L[:-5].strip()
            else:
                L2 = L
            return (V == L) or (V2 == L) or (V == L2) or (V2 == L2)
        if group == "test_status" and label == "...TH":
            return V.endswith("TH") and V not in {"1ST", "2ND", "3RD"}
        return V == L

    label_value_map = {}
    for group, labels in label_groups.items():
        value = data.get(group, None)
        if value is None or (isinstance(value, str) and not value.strip()):
            for label in labels:
                label_value_map[label] = False
        else:
            for label in labels:
                label_value_map[label] = _eq_relaxed(label, str(value), group)

    # N/A flags
    for field in ["sample_description", "item_code", "supplier", "subcon"]:
        val = str(data.get(field, "")).strip().upper()
        label_value_map[f"{field.upper()} N/A"] = (val == "N/A")

    return label_value_map

# =========================
# Tick checkbox theo ĐÚNG NHÃN
# =========================
def _label_regex(label: str) -> re.Pattern:
    """
    Tạo regex khớp '☐/☑' ngay TRƯỚC nhãn tương ứng, không phân biệt hoa thường/khoảng trắng.
    Ví dụ: 'CONSTRUCTION TEST' -> r'(☐|☑)\s*C\s*O\s*N...'
    """
    # chuẩn hóa nhãn: bỏ các dấu '_' '.' '-', co khoảng trắng
    cleaned = re.sub(r'[_\.\-]+', ' ', (label or '').strip())
    parts = [p for p in cleaned.split() if p]
    pattern = r'(☐|☑)\s*' + r'\s*'.join(re.escape(p) for p in parts)
    return re.compile(pattern, flags=re.IGNORECASE)

def tick_unicode_checkbox_by_label(doc: Document, label_value_map):
    """
    Duyệt paragraph + table cell, chỉ thay dấu checkbox đứng NGAY TRƯỚC nhãn khớp.
    Không 'replace toàn bộ' -> không dẫm chân các checkbox khác trong cùng ô/đoạn.
    """
    compiled = [( _label_regex(label), bool(value) ) for label, value in label_value_map.items()]

    def toggle_text(txt: str) -> str:
        if not txt or ('☐' not in txt and '☑' not in txt):
            return txt
        for pat, value in compiled:
            def _repl(m):
                # m.group(0) bắt đầu bằng checkbox → thay đúng checkbox đó
                return ('☑' if value else '☐') + m.group(0)[1:]
            txt = pat.sub(_repl, txt)
        return txt

    for p in doc.paragraphs:
        t = p.text
        if ("☐" in t) or ("☑" in t):
            p.text = toggle_text(t)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text
                if ("☐" in t) or ("☑" in t):
                    cell.text = toggle_text(t)

# =========================
# Convert PDF (best-effort)
# =========================
def try_convert_to_pdf(docx_path, pdf_path):
    """
    DOCX -> PDF bằng docx2pdf (Word). Không raise để không chặn approve.
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
        from docx2pdf import convert
        convert(docx_path, pdf_path)
    except Exception as e:
        import traceback
        print("Không thể convert PDF:", e)
        traceback.print_exc()

# =========================
# Per-report lock + atomic save
# =========================
def _lock_path_for(report_no: str) -> str:
    return os.path.join(tempfile.gettempdir(), f"tfr_{report_no}.lock")

def _acquire_lock(path: str, timeout=30):
    t0 = time.time()
    fd = None
    while True:
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
            return fd
        except FileExistsError:
            if time.time() - t0 > timeout:
                try: os.unlink(path)
                except: pass
            time.sleep(0.05)
        except Exception:
            time.sleep(0.05)

def _release_lock(fd, path: str):
    try: os.close(fd)
    except: pass
    try: os.unlink(path)
    except: pass

def _atomic_save_docx(doc: Document, out_path: str):
    tmp = f"{out_path}.tmp-{uuid.uuid4().hex}"
    doc.save(tmp)
    os.replace(tmp, out_path)  # atomic trên cùng filesystem

# =========================
# Fill DOCX + xuất PDF
# =========================
def fill_docx_and_export_pdf(data, fixed_report_no=None):
    """
    - Nếu fixed_report_no có giá trị => dùng y nguyên.
    - Ngược lại => lấy report_no trống đầu tiên theo C..X đều trống.
    - Điền vào template, tick checkbox, atomic save DOCX, cố gắng convert PDF.
    """
    if fixed_report_no and str(fixed_report_no).strip():
        report_no = str(fixed_report_no).strip()
    else:
        report_no = get_first_empty_report_all_blank(local_main)
        if not report_no:
            raise Exception("Không còn mã report trống trong Excel.")

    # copy để không làm bẩn dict gốc
    data = dict(data)
    data["report_no"] = report_no

    # 1) Mở template
    doc = Document(WORD_TEMPLATE)

    # 2) Mapping điền ô bảng (label cột trái → value cột phải)
    mapping = {
        "requestor": "requestor",
        "department": "department",
        "requested date": "request_date",
        "lab test report no.": "report_no",
        "sample description": "sample_description",
        "item code": "item_code",
        "quantity": "quantity",
        "supplier": "supplier",
        "subcon": "subcon",
        "test group": "test_group",
        "test status": "test_status",
        "furniture testing": "furniture_testing",
        "estimated completed date": "estimated_completion_date",
    }
    remark = data.get("remark", "")
    remark_written = False

    # 3) Điền các ô theo mapping
    for table in doc.tables:
        nrows = len(table.rows)
        ncols = len(table.columns)
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                label = (
                    cell.text.strip().lower()
                    .replace("(mã item)", "")
                    .replace("(mã material)", "")
                    .replace("*", "")
                )
                # remark: ghi xuống ô bên dưới nếu trống
                if not remark_written and ("other tests/instructions" in label or "remark" in label) and remark:
                    if i + 1 < nrows:
                        below_cell = table.rows[i + 1].cells[j]
                        if not (below_cell.text or "").strip():
                            below_cell.text = str(remark)
                            remark_written = True
                            continue
                # employee id / msnv
                if ("emp id" in label or "msnv" in label) and data.get("employee_id", ""):
                    if j + 1 < ncols:
                        target_cell = row.cells[j + 1]
                        if not (target_cell.text or "").strip():
                            target_cell.text = str(data["employee_id"])
                            continue
                # mapping fields
                for map_label, key in mapping.items():
                    if map_label in ["remark", "employee id"]:
                        continue
                    if map_label in label and key in data and str(data[key]).strip() != "":
                        if j + 1 < ncols:
                            target_cell = row.cells[j + 1]
                            # luôn cho phép set "lab test report no." để ghi đè
                            if (target_cell.text or "").strip() == "" or "lab test report no." in label:
                                target_cell.text = str(data[key])

    # 4) Tick checkbox theo nhãn (mới: không replace toàn cục)
    label_value_map = build_label_value_map(data)
    tick_unicode_checkbox_by_label(doc, label_value_map)

    # 5) Đảm bảo thư mục output
    if not os.path.exists(PDF_OUTPUT_FOLDER):
        os.makedirs(PDF_OUTPUT_FOLDER)
    output_docx = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.docx")
    output_pdf  = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.pdf")

    # 6) Ghi atomic + convert PDF (best-effort) dưới lock theo report_no
    lock_path = _lock_path_for(report_no)
    fd = _acquire_lock(lock_path, timeout=30)
    try:
        _atomic_save_docx(doc, output_docx)
    finally:
        _release_lock(fd, lock_path)

    try_convert_to_pdf(output_docx, output_pdf)

    return output_docx, output_pdf, report_no

# =========================
# API cho app.py
# =========================
def approve_request_fill_docx_pdf(req):
    """
    Hàm wrapper để app.allocate_unique_report_no() gọi:
      - Nếu req chứa 'report_no' -> điền đúng số này (validate ở bên app).
      - Nếu không có -> tự tìm dòng C..X trống để lấy report_no.
    Trả về: (pdf_path_hoặc_docx_path, report_no)
    """
    fixed = (req.get("report_no") or "").strip()
    # Chuẩn hóa tên trường ETD cho Word nếu có
    if "etd" in req and not req.get("estimated_completion_date"):
        req = dict(req)
        req["estimated_completion_date"] = req.get("etd")

    if fixed:
        out_docx, out_pdf, report_no = fill_docx_and_export_pdf(req, fixed_report_no=fixed)
    else:
        out_docx, out_pdf, report_no = fill_docx_and_export_pdf(req, fixed_report_no=None)

    # Ưu tiên PDF nếu có
    if os.path.exists(out_pdf):
        return out_pdf, report_no
    return out_docx, report_no
