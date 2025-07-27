import pandas as pd
import re
from config import local_main
from openpyxl import load_workbook

# Hàm tìm chỉ số cột theo tên cột
def get_col_idx(ws, target):
    norm_target = normalize_colname(target)
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v and normalize_colname(v) == norm_target:
            return col
    return None

def ensure_column(ws, col_name):
    header = [cell.value for cell in ws[1]]
    if col_name not in header:
        ws.cell(row=1, column=ws.max_column+1, value=col_name)
    return [cell.value for cell in ws[1]].index(col_name) + 1

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

def clean_col(s):
    s = str(s).lower()
    s = re.sub(r'[^a-z0-9#]+', '', s)
    return s

QAD_DF = pd.read_excel(local_main)
QAD_DF.columns = [clean_col(c) for c in QAD_DF.columns]

def get_item_code(report):
    row = QAD_DF.loc[QAD_DF['report#'] == str(report)]
    if not row.empty:
        return str(row.iloc[0]['item#'])
    return ""

def normalize_colname(s):
    return re.sub(r'[\s/\.:\n]+', '', str(s).strip().lower())

def write_tfr_to_excel(excel_path, report_no, request):
    wb = load_workbook(excel_path)
    ws = wb.active

    # Tìm header
    headers = {}
    for col in range(1, ws.max_column + 1):
        name = ws.cell(row=1, column=col).value
        if name:
            clean = (
                str(name)
                .strip()
                .replace('\n', ' ')
                .replace('/', ' ')
                .replace('.', '')
                .replace('#', '')
                .lower()
            )
            clean = " ".join(clean.split())
            headers[clean] = col

    # Tìm cột REPORT NO
    report_col = None
    for k in headers:
        if "report" in k and "no" in k:
            report_col = headers[k]
            break
    if not report_col:
        report_col = 1  # fallback cột đầu tiên nếu không tìm được!

    # Tìm dòng có report_no (so sánh số, chuỗi, bỏ khoảng trắng, in hoa)
    row_idx = None
    for row in range(2, ws.max_row + 1):
        cell_val = ws.cell(row=row, column=report_col).value
        if cell_val is not None:
            # So sánh full string
            if str(cell_val).strip().upper() == str(report_no).strip().upper():
                row_idx = row
                break
            # Nếu mã là dạng 25-4822, kiểm tra phần số sau
            try:
                report_no_num = int(str(report_no).replace('25-', '').replace('-', '').strip())
                cell_num = int(str(cell_val).replace('25-', '').replace('-', '').strip())
                if report_no_num == cell_num:
                    row_idx = row
                    break
            except:
                pass

    if row_idx is None:
        raise Exception(f"Không tìm thấy mã report {report_no} trong file excel!")

    def to_upper(val):
        if isinstance(val, str):
            return val.upper()
        return val

    fields_map = [
        (["item"], to_upper(request.get("item_code", ""))),
        (["type of"], to_upper(", ".join(request.get("test_groups", [])))),
        (["item name", "description"], to_upper(request.get("sample_description", ""))),
        (["furniture testing"], to_upper(request.get("furniture_testing", ""))),
        (["submitter in", "submitter in charge"], to_upper(request.get("requestor", ""))),
        (["submitted dept"], to_upper(request.get("department", ""))),
        (["remark"], to_upper(request.get("test_status", ""))),
    ]

    for keys, val in fields_map:
        for k in headers:
            if all(word in k for word in [w.lower() for w in keys]):
                ws.cell(row=row_idx, column=headers[k]).value = val
                break

    wb.save(excel_path)
