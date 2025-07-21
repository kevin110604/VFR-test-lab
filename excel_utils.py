import pandas as pd
import re
from config import local_main

def get_col_idx(ws, target):
    for col in range(1, ws.max_column+1):
        v = ws.cell(row=1, column=col).value
        if v and target.strip().lower() in str(v).strip().lower():
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
    # Đưa về chữ thường, loại bỏ khoảng trắng, bỏ ký tự đặc biệt trừ dấu #
    s = str(s).lower()
    s = re.sub(r'[^a-z0-9#]+', '', s)  # giữ lại chữ, số, dấu #
    return s

QAD_DF = pd.read_excel(local_main)
QAD_DF.columns = [clean_col(c) for c in QAD_DF.columns]

def get_item_code(report):
    # Chuẩn hóa tên cột ở đây
    # 'report#' (không khoảng trắng, không phân biệt hoa/thường)
    row = QAD_DF.loc[QAD_DF['report#'] == str(report)]
    if not row.empty:
        return str(row.iloc[0]['item#'])  # item#
    return ""
