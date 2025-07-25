import zipfile
from lxml import etree
import shutil
import os
from openpyxl import load_workbook
from docx import Document
from docx2pdf import convert
from config import local_main

WORD_TEMPLATE = "FORM-QAD-011-TEST REQUEST FORM (TRF).docx"
PDF_OUTPUT_FOLDER = os.path.join("static", "TFR")

def get_first_empty_report_all_blank(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        report_no = str(row[0]).strip() if row[0] else None
        if not report_no:
            continue
        is_blank = True
        for idx, cell in enumerate(row):
            if idx == 0 or idx == 20:
                continue
            if cell not in (None, "", " "):
                is_blank = False
                break
        if is_blank:
            return report_no
    return None

def build_label_value_map(data):
    label_groups = {
        "sample_type": ["MATERIAL", "CARCASS", "FINISHED ITEM", "OTHERS"],
        "test_status": ["1ST", "2ND", "3RD", "...TH"],
        "furniture_testing": ["INDOOR", "OUTDOOR"],
        "test_groups": [
            "CONSTRUCTION TEST",
            "PACKAGING TEST (TRANSIT TEST)",
            "MATERIAL AND FINISHING TEST"
        ]
    }
    label_value_map = {}
    for group, labels in label_groups.items():
        value = data.get(group, None)
        if value is None or (isinstance(value, str) and not value.strip()) or (isinstance(value, list) and len(value) == 0):
            for label in labels:
                label_value_map[label] = False
        else:
            if not isinstance(value, list):
                value_list = [value]
            else:
                value_list = value
            value_list = [str(v).strip().upper() for v in value_list]
            for label in labels:
                if group == "test_status" and label == "...TH":
                    label_value_map[label] = any("NTH" in v for v in value_list)
                else:
                    label_value_map[label] = (label in value_list)
    return label_value_map

def tick_unicode_checkbox_by_label(docx_path, out_path, label_value_map):
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
            if '☐' in t.text or '☑' in t.text:
                # Lấy label ngay sau tick (ghép 1-2 block nếu cần)
                label_parts = []
                # Lấy block ngay sau tick
                if i+1 < len(texts):
                    label_parts.append(texts[i+1].text or "")
                # Có thể ghép thêm block nữa nếu label bị tách nhỏ
                if i+2 < len(texts):
                    label_parts.append(texts[i+2].text or "")
                label = ''.join(label_parts).strip().replace(" ", "").replace("\n", "").upper()
                # So khớp label duy nhất với từng key, nếu đúng thì chỉ tick đúng 1 lần
                for label_key, value in label_value_map.items():
                    key_norm = label_key.replace(" ", "").replace("_", "").replace(".", "").upper()
                    if label.startswith(key_norm):
                        old = t.text
                        if value:
                            t.text = t.text.replace('☐', '☑')
                        else:
                            t.text = t.text.replace('☑', '☐')
                        print(f"[DEBUG] Tick {label_key}: {old} → {t.text} (tick next to: '{label}')")
                        break  # Đã tick đúng, không tick lại tick này nữa
            i += 1

    tree.write(xml_path, xml_declaration=True, encoding='utf-8', standalone=True)
    shutil.make_archive("output_docx", 'zip', temp_dir)
    shutil.move("output_docx.zip", out_path)
    shutil.rmtree(temp_dir)

def try_convert_to_pdf(docx_path, pdf_path):
    try:
        import pythoncom
        pythoncom.CoInitialize()  # <--- Thêm dòng này!
        from docx2pdf import convert
        convert(docx_path, pdf_path)
    except Exception as e:
        import traceback
        print("Không thể convert PDF:", e)
        traceback.print_exc()
        
def fill_docx_and_export_pdf(data):
    report_no = get_first_empty_report_all_blank(local_main)
    if not report_no:
        raise Exception("Không còn mã report trống trong Excel.")
    data = dict(data)
    data["report_no"] = report_no

    doc = Document(WORD_TEMPLATE)
    mapping = {
        "requestor": "requestor",
        "department": "department",
        "title": "title",
        "requested date": "request_date",
        "estimated completed date": "estimated_completion_date",
        "lab test report no.": "report_no",
        "sample description": "sample_description",
        "item code": "item_code",
        "material code": "material_code",
        "quantity": "quantity",
        "supplier": "supplier",
        "subcon": "subcon",
        "dimension": "dimension",
        "sales code": "sales_code",
        "weight": "weight",
    }
    for table in doc.tables:
        nrows = len(table.rows)
        ncols = len(table.columns)
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                label = cell.text.strip().lower().replace("(mã item)", "").replace("(mã material)", "").replace("*", "")
                for map_label, key in mapping.items():
                    if map_label in label and key in data and str(data[key]).strip() != "":
                        if j+1 < ncols:
                            target_cell = row.cells[j+1]
                            if target_cell.text.strip() == "" or "lab test report no." in label:
                                target_cell.text = str(data[key])

    if not os.path.exists(PDF_OUTPUT_FOLDER):
        os.makedirs(PDF_OUTPUT_FOLDER)
    output_docx = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.docx")
    doc.save(output_docx)

    label_value_map = build_label_value_map(data)
    tick_unicode_checkbox_by_label(output_docx, output_docx, label_value_map)

    output_pdf = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.pdf")
    try_convert_to_pdf(output_docx, output_pdf)

    # Đường dẫn trả về để dùng cho url_for('static', ...)
    pdf_path = f"TFR/{report_no}.pdf"
    return pdf_path, report_no

def approve_request_fill_docx_pdf(data_dict):
    return fill_docx_and_export_pdf(data_dict)
