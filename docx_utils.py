# docx_utils.py
import os
import time
import uuid
import tempfile
from openpyxl import load_workbook
from docx import Document
from config import local_main

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
            report_no = ws.cell(row=row, column=2).value  # B
            if report_no is not None and str(report_no).strip() != "":
                return str(report_no).strip()
    return None

# =========================
# Checkbox mapping cho DOCX
# =========================
def build_label_value_map(data):
    label_groups = {
        "test_group": [
            "MATERIAL TEST",
            "FINISHING TEST",
            "CONSTRUCTION TEST",
            "TRANSIT TEST",
            "ENVIRONMENTAL TEST",
        ],
        "test_status": ["1ST", "2ND", "3RD", "...TH"],  # nth → ...th
        "furniture_testing": ["INDOOR", "OUTDOOR"],
        "sample_return": ["YES", "NO"]
    }
    label_value_map = {}
    for group, labels in label_groups.items():
        value = data.get(group, None)
        if value is None or (isinstance(value, str) and not value.strip()):
            for label in labels:
                label_value_map[label] = False
        else:
            vstr = str(value).strip().upper()
            for label in labels:
                if group == "test_status" and label == "...TH":
                    label_value_map[label] = (vstr.endswith("TH") and vstr not in ["1ST", "2ND", "3RD"])
                else:
                    label_value_map[label] = (label == vstr)
    # N/A flags
    for field in ["sample_description", "item_code", "supplier", "subcon"]:
        val = str(data.get(field, "")).strip().upper()
        label_value_map[f"{field.upper()} N/A"] = (val == "N/A")
    return label_value_map

def tick_unicode_checkbox_by_label(doc: Document, label_value_map):
    """
    Không unzip/nén lại. Duyệt paragraph + table cell, thay '☐'/'☑' theo label_value_map.
    """
    def toggle_text(txt: str, label_map):
        txt_norm = txt.replace(" ", "").replace("\n", "").upper()
        for label_key, value in label_map.items():
            key_norm = label_key.replace(" ", "").replace("_", "").replace(".", "").upper()
            if key_norm in txt_norm:
                txt = txt.replace("☐", "☑") if value else txt.replace("☑", "☐")
        return txt

    for p in doc.paragraphs:
        t = p.text
        if ("☐" in t) or ("☑" in t):
            p.text = toggle_text(t, label_value_map)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text
                if ("☐" in t) or ("☑" in t):
                    cell.text = toggle_text(t, label_value_map)

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
# Per-report lock + atomic save helpers
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
                # stale -> cố gắng xóa rồi lấy
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
    # Thay thế atomic để không bao giờ đọc file dở dang
    os.replace(tmp, out_path)

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

    data = dict(data)
    data["report_no"] = report_no

    doc = Document(WORD_TEMPLATE)
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

    # Điền các ô trong bảng theo mapping (label ở cột trái, value ở cột phải)
    for table in doc.tables:
        nrows = len(table.rows)
        ncols = len(table.columns)
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                label = cell.text.strip().lower().replace("(mã item)", "").replace("(mã material)", "").replace("*", "")
                # remark
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

    # Tick checkbox theo label map (trên đối tượng doc, không unzip)
    label_value_map = build_label_value_map(data)
    tick_unicode_checkbox_by_label(doc, label_value_map)

    # Bảo đảm thư mục output
    if not os.path.exists(PDF_OUTPUT_FOLDER):
        os.makedirs(PDF_OUTPUT_FOLDER)
    output_docx = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.docx")
    output_pdf  = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.pdf")

    # Khóa ngắn theo report_no để tránh 2 luồng cùng lúc ghi & convert
    lock_path = _lock_path_for(report_no)
    fd = _acquire_lock(lock_path, timeout=30)
    try:
        # Atomic save DOCX
        _atomic_save_docx(doc, output_docx)
        # Convert PDF (best-effort)
        try_convert_to_pdf(output_docx, output_pdf)
    finally:
        _release_lock(fd, lock_path)

    pdf_path = f"TFR/{report_no}.pdf"
    return pdf_path, report_no

# =========================
# Entry cho "Approve"
# =========================
def approve_request_fill_docx_pdf(data_dict):
    """
    - Nếu data_dict có report_no thì dùng đúng report_no đó.
    - Nếu chưa có -> chọn report_no trống theo C..X.
    - Trả về (pdf_path, report_no).
    """
    report_no = str(data_dict.get("report_no", "")).strip()
    if not report_no:
        report_no = get_first_empty_report_all_blank(local_main)
        if not report_no:
            raise Exception("Không còn mã report trống trong Excel.")
        data_dict["report_no"] = report_no

    return fill_docx_and_export_pdf(data_dict, fixed_report_no=report_no)
