import pandas as pd
import re
from config import local_main
from openpyxl import load_workbook

# Hàm tìm chỉ số cột theo tên cột
def get_col_idx(ws, target):
    for col in range(1, ws.max_column+1):
        v = ws.cell(row=1, column=col).value
        if v and target.strip().lower() in str(v).strip().lower():
            return col
    return None

# Đảm bảo tồn tại cột, nếu chưa có thì thêm mới vào cuối
def ensure_column(ws, col_name):
    header = [cell.value for cell in ws[1]]
    if col_name not in header:
        ws.cell(row=1, column=ws.max_column+1, value=col_name)
    return [cell.value for cell in ws[1]].index(col_name) + 1

# Copy nguyên một dòng có cả style giữa 2 sheet
def copy_row_with_style(ws_from, ws_to, row_idx, to_row=None):
    if to_row is None:
        to_row = ws_to.max_row + 1
    for col in range(1, ws_from.max_column+1):
        c1 = ws_from.cell(row=row_idx, column=col)
        c2 = ws_to.cell(row=to_row, column=col)
        c2.value = c1.value
        try:
            c2.font = c1.font.copy() if c1.font else None
            c2.border = c1.border.copy() if c1.border else None
            c2.fill = c1.fill.copy() if c1.fill else None
            c2.number_format = c1.number_format
            c2.protection = c1.protection.copy() if c1.protection else None
            c2.alignment = c1.alignment.copy() if c1.alignment else None
        except Exception:
            pass

# Kiểm tra có hình tại cell không
def is_img_at_cell(img, row, col):
    try:
        anchor = getattr(img, "anchor", None)
        if hasattr(anchor, "_from"):
            return (anchor._from.row + 1 == row) and (anchor._from.col + 1 == col)
        if hasattr(anchor, "cell"):
            img_row, img_col = anchor.cell.row + 1, anchor.cell.col_idx
            return (img_row == row) and (img_col == col)
    except Exception:
        return False
    return False

# Chuẩn hóa tên cột: chữ thường, loại bỏ ký tự đặc biệt trừ dấu #
def clean_col(s):
    s = str(s).lower()
    s = re.sub(r'[^a-z0-9#]+', '', s)  # giữ lại chữ, số, dấu #
    return s

# Đọc file QAD_EXCEL thành DataFrame pandas (đã chuẩn hóa tên cột)
QAD_DF = pd.read_excel(local_main)
QAD_DF.columns = [clean_col(c) for c in QAD_DF.columns]

# Lấy item code theo report code
def get_item_code(report):
    # report# (không khoảng trắng, không phân biệt hoa/thường)
    row = QAD_DF.loc[QAD_DF['report#'] == str(report)]
    if not row.empty:
        return str(row.iloc[0]['item#'])  # item#
    return ""

from openpyxl import load_workbook

from openpyxl import load_workbook
import re

def normalize_colname(s):
    # Chuẩn hóa tên cột để dò cho "khôn": xóa mọi khoảng trắng, dấu /, dấu chấm, in thường
    return re.sub(r'[\s/\.:\n]+', '', str(s).strip().lower())

def write_tfr_to_excel(excel_path, report_no, request):
    wb = load_workbook(excel_path)
    ws = wb.active

    # Đọc headers, chuẩn hóa để dò cột
    headers = {}
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        name = cell.value
        if name:
            headers[normalize_colname(name)] = col

    # Tìm cột "report no"
    report_col = None
    for k in headers:
        if "report" in k and "no" in k:
            report_col = headers[k]
            break

    # Tìm dòng theo report_no, hoặc thêm mới nếu chưa có
    row_idx = None
    if report_col:
        for row in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=report_col).value
            if str(cell_val).strip() == str(report_no):
                row_idx = row
                break
        if row_idx is None:
            row_idx = ws.max_row + 1
            ws.cell(row=row_idx, column=report_col).value = report_no
    else:
        row_idx = ws.max_row + 1

    # CHỈ value ghi vào là in hoa hết, header để nguyên!
    def to_upper(val):
        return str(val).upper() if val else ""

    # Map trường bạn yêu cầu:
    field_map = [
        # (các từ khóa cần match, value cần ghi vào, comment)
        (["item#"], to_upper(request.get("item_code", ""))),  # Item#
        (["typeof"], to_upper(", ".join(request.get("test_groups", [])))),  # Type of
        (["itemname", "description"], to_upper(request.get("sample_description", ""))),  # Item name/Description
        (["furnituretesting"], to_upper(request.get("furniture_testing", ""))),  # Furniture testing
        (["submitterinchange"], to_upper(request.get("requestor", ""))),  # Submitter in change
        (["submitteddept"], to_upper(request.get("department", ""))),  # Submitted dept.
        (["remark"], to_upper(request.get("test_status", ""))),  # Remark
    ]

    # Ghi lần lượt từng trường vào đúng cột, đúng dòng
    for keys, value in field_map:
        col_found = None
        for header, col in headers.items():
            for key in keys:
                if key in header:
                    col_found = col
                    break
            if col_found:
                break
        if col_found:
            ws.cell(row=row_idx, column=col_found).value = value

    wb.save(excel_path)
