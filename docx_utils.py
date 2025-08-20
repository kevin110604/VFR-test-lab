# docx_utils.py
import os
import shutil
import zipfile
from lxml import etree
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
    """
    Trả về (is_blank, normalized_str).
    Trống nếu: None hoặc chuỗi chỉ còn whitespace / '-' / '—' sau khi loại NBSP/zero-width/CRLF/TAB.
    0/False/ngày tháng -> KHÔNG trống.
    """
    if v is None:
        return True, ""
    if isinstance(v, str):
        s = (
            v.replace("\u00A0", "")   # NBSP
             .replace("\u200B", "")   # zero-width
             .replace("\r", "")
             .replace("\n", "")
             .replace("\t", "")
             .strip()
        )
        return (s in BLANK_TOKENS), s
    return False, str(v)

def get_first_empty_report_all_blank(excel_path):
    """
    Trả về report_no của dòng đầu tiên mà các ô C..X (3..24) đều TRỐNG.
    - B(2) = report_no (được trả về)
    - Y(25) = QR/link (bỏ qua khi xét trống)
    """
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
    # N/A flags cho 4 field text
    for field in ["sample_description", "item_code", "supplier", "subcon"]:
        val = str(data.get(field, "")).strip().upper()
        label_value_map[f"{field.upper()} N/A"] = (val == "N/A")
    return label_value_map

def tick_unicode_checkbox_by_label(docx_path, out_path, label_value_map):
    """
    Mở file DOCX như zip, sửa document.xml:
    - Nếu chữ ngay sau ký tự '☐'/'☑' khớp key trong label_value_map (so khớp prefix đã normalize),
      thay checkbox theo True/False.
    """
    temp_dir = 'temp_unzip_docx'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    xml_path = os.path.join(temp_dir, 'word', 'document.xml')
    tree = etree.parse(xml_path)
    root = tree.getroot()
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for cell in root.xpath(".//w:tc", namespaces=ns):
        texts = cell.xpath(".//w:t", namespaces=ns)
        i = 0
        while i < len(texts):
            t = texts[i]
            txt = t.text or ""
            if ('☐' in txt) or ('☑' in txt):
                # gom nhãn kế sau checkbox (thường split thành nhiều w:t)
                label_parts = []
                if i + 1 < len(texts):
                    label_parts.append(texts[i + 1].text or "")
                if i + 2 < len(texts):
                    label_parts.append(texts[i + 2].text or "")
                label = ''.join(label_parts).strip().replace(" ", "").replace("\n", "").upper()

                for label_key, value in label_value_map.items():
                    key_norm = label_key.replace(" ", "").replace("_", "").replace(".", "").upper()
                    if label.startswith(key_norm):
                        if value:
                            t.text = txt.replace('☐', '☑')
                        else:
                            t.text = txt.replace('☑', '☐')
                        break
            i += 1

    tree.write(xml_path, xml_declaration=True, encoding='utf-8', standalone=True)
    shutil.make_archive("output_docx", 'zip', temp_dir)
    shutil.move("output_docx.zip", out_path)
    shutil.rmtree(temp_dir)

# =========================
# Convert PDF (best-effort)
# =========================
def try_convert_to_pdf(docx_path, pdf_path):
    """
    Cố gắng convert DOCX -> PDF bằng docx2pdf (chỉ hoạt động trên Windows với Word).
    Không raise lỗi để không chặn luồng approve.
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
# Fill DOCX + xuất PDF
# =========================
def fill_docx_and_export_pdf(data, fixed_report_no=None):
    """
    - Nếu fixed_report_no có giá trị => dùng y nguyên.
    - Ngược lại => lấy report_no trống đầu tiên theo C..X đều trống.
    - Điền vào template, tick checkbox, lưu DOCX, cố gắng convert PDF.
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

    # Lưu DOCX
    if not os.path.exists(PDF_OUTPUT_FOLDER):
        os.makedirs(PDF_OUTPUT_FOLDER)
    output_docx = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.docx")
    doc.save(output_docx)

    # Tick checkbox theo label map
    label_value_map = build_label_value_map(data)
    tick_unicode_checkbox_by_label(output_docx, output_docx, label_value_map)

    # Convert PDF (best-effort)
    output_pdf = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.pdf")
    try_convert_to_pdf(output_docx, output_pdf)

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
