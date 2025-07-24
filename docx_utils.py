import zipfile
from lxml import etree
import shutil
import os
from openpyxl import load_workbook
from docx import Document
from docx2pdf import convert
from config import local_main

WORD_TEMPLATE = "FORM-QAD-011-TEST REQUEST FORM (TRF).docx"
PDF_OUTPUT_FOLDER = "TFR"

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

def get_label_to_tick(data):
    # Tập hợp tất cả các lựa chọn có thể của từng nhóm
    label_groups = {
        "sample_type": ["Material", "Carcass", "Finished Item", "Others"],
        "test_status": ["1st", "2nd", "3rd", "...th"],
        "furniture_testing": ["Indoor", "Outdoor"],
        "test_groups": [
            "CONSTRUCTION TEST",
            "PACKAGING TEST (TRANSIT TEST)",
            "MATERIAL AND FINISHING TEST"
        ]
    }
    label_to_tick = {}

    # Tick từng nhóm, nếu có dữ liệu thì tick đúng lựa chọn, các lựa chọn khác tick False
    for group, labels in label_groups.items():
        val = data.get(group, None)
        if val is None or (isinstance(val, str) and not val.strip()) or (isinstance(val, list) and len(val) == 0):
            # Không điền -> tick False tất cả
            for label in labels:
                label_to_tick[label] = False
        else:
            if not isinstance(val, list):
                val_list = [val]
            else:
                val_list = val
            val_list = [str(v).strip().lower() for v in val_list]
            for label in labels:
                if group == "test_status":
                    if label == "...th":
                        label_to_tick[label] = any("nth" in v for v in val_list)
                    else:
                        label_to_tick[label] = (label.lower() in val_list)
                else:
                    label_to_tick[label] = (label.lower() in val_list)

    return label_to_tick

def tick_checkbox_by_label(docx_path, out_path, label_to_tick):
    temp_dir = 'temp_unzip_docx'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    xml_path = os.path.join(temp_dir, 'word', 'document.xml')
    tree = etree.parse(xml_path)
    root = tree.getroot()
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for checkbox in root.xpath(".//w:checkBox", namespaces=ns):
        node = checkbox.getparent()
        label_full = ""
        # Tìm label gần checkbox (lấy ancestor có text)
        for ancestor in [node] + list(node.iterancestors()):
            text_list = ancestor.xpath(".//w:t/text()", namespaces=ns)
            if text_list:
                label_full = " ".join([t.strip() for t in text_list if t.strip()])
                for label_key in label_to_tick:
                    # Fuzzy match (label_key nằm trong text)
                    if label_key.lower() in label_full.lower():
                        checked = checkbox.find(".//w:checked", namespaces=ns)
                        if checked is not None:
                            checked.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"] = "1" if label_to_tick[label_key] else "0"
                            print(f"[DEBUG] Tick '{label_key}': {label_to_tick[label_key]} (found in: '{label_full}')")
                        break
                break

    tree.write(xml_path, xml_declaration=True, encoding='utf-8', standalone=True)
    shutil.make_archive("output_docx", 'zip', temp_dir)
    shutil.move("output_docx.zip", out_path)
    shutil.rmtree(temp_dir)

def try_convert_to_pdf(docx_path, pdf_path):
    try:
        convert(docx_path, pdf_path)
    except Exception as e:
        print("Không thể convert PDF, có thể do chưa cài MS Word hoặc chạy trên Linux/Mac. Bỏ qua PDF. Lý do:", e)

def fill_docx_and_export_pdf(data):
    report_no = get_first_empty_report_all_blank(local_main)
    if not report_no:
        raise Exception("Không còn mã report trống trong Excel.")
    data = dict(data)
    data["report_no"] = report_no

    # Ghi text như cũ
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

    # Save file docx (chưa tick)
    if not os.path.exists(PDF_OUTPUT_FOLDER):
        os.makedirs(PDF_OUTPUT_FOLDER)
    output_docx = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.docx")
    doc.save(output_docx)

    # Tick theo label (auto từng nhóm)
    label_to_tick = get_label_to_tick(data)
    tick_checkbox_by_label(output_docx, output_docx, label_to_tick)

    # Convert pdf nếu cần
    output_pdf = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.pdf")
    try_convert_to_pdf(output_docx, output_pdf)
    return output_pdf, report_no

def approve_request_fill_docx_pdf(data_dict):
    return fill_docx_and_export_pdf(data_dict)
