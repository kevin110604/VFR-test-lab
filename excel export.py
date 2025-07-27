import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import io
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from datetime import datetime

# ==== Cấu hình SharePoint và tên file xuất ====
site_url = "https://jonathancharles.sharepoint.com/sites/TESTLAB-VFR9"
username = "tan_qa@vfr.net.vn"
password = "qaz@Tat@123"
relative_url = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/QAD-Outstanding list-2025.xlsx"
excel_file_out = "ds sản phẩm test với qr.xlsx"

# ==== Đọc file gốc từ SharePoint ====
ctx_auth = AuthenticationContext(site_url)
if not ctx_auth.acquire_token_for_user(username, password):
    raise Exception("Không kết nối được SharePoint!")
ctx = ClientContext(site_url, ctx_auth)

download = io.BytesIO()
file = ctx.web.get_file_by_server_relative_url(relative_url)
file.download(download).execute_query()
download.seek(0)

excel_file = pd.ExcelFile(download)
sheet_name = None
for name in excel_file.sheet_names:
    if name.strip().lower() == "ol list":
        sheet_name = name
        break
if sheet_name is None:
    raise Exception("Không tìm thấy sheet 'OL list'!")
df = excel_file.parse(sheet_name)

# ==== Chọn các cột B-Q (2-17), U,V (21-22) ====
cols_selected = list(df.columns[1:17]) + list(df.columns[20:22])

# ==== Tìm các cột "type of", "status", "report" ====
def find_col(keywords):
    for col in df.columns:
        if all(k in col.lower() for k in keywords):
            return col
    for col in df.columns:
        if any(k in col.lower() for k in keywords):
            return col
    return None

type_of_cols = [col for col in df.columns if "type of" in col.lower()]
status_col = find_col(["status"])
report_col = find_col(["report", "#"])

def norm(s):
    return (str(s).strip().lower() if not pd.isnull(s) else "")

# ==== Lọc dòng: loại bỏ outsource ====
rows = []
for idx, row in df.iterrows():
    skip = False
    for col in type_of_cols:
        val = norm(row[col])
        if "outsource-mts" in val or "outsource-sgs" in val:
            skip = True
            break
    if skip:
        continue
    rows.append(idx)
df_out = df.iloc[rows].copy()
df_out = df_out[cols_selected]

# ==== Lọc theo Report ID >= 4500 ====
report_col_main = report_col if report_col in df_out.columns else df_out.columns[0]
df_out["__report_num__"] = pd.to_numeric(df_out[report_col_main].astype(str).str.extract(r'(\d+)$')[0], errors="coerce")
df_out = df_out[df_out["__report_num__"] >= 4500]
df_out.drop(columns="__report_num__", inplace=True)

# ==== Thêm cột bổ sung ====
for col in ["Test Date", "Complete Date", "Rating"]:
    df_out[col] = ""

# ==== Tạo cột QR Code link từ mã 25-xxxx ====
qr_url_dict = {}
for idx, row in df_out.iterrows():
    report_raw = str(row[report_col_main]).strip()
    if report_raw and report_raw.lower() != "nan":
        url = f"http://103.77.166.187:8080/update?report={report_raw}"
        qr_url_dict[report_raw] = url
df_out["QR Code"] = df_out[report_col_main].astype(str).map(qr_url_dict).fillna("")

# ==== CẮT GIỜ, CHỈ LẤY NGÀY-THÁNG-NĂM ====
def only_date(val):
    if pd.isnull(val) or not str(val).strip():
        return ""
    try:
        # Nếu là datetime
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
        # Nếu là chuỗi có định dạng ngày-giờ
        s = str(val)
        if "-" in s and len(s.split(" ")[0]) == 10:
            return datetime.strptime(s.split(" ")[0], "%Y-%m-%d").strftime("%d/%m/%Y")
        if "/" in s and len(s.split(" ")[0]) == 10:
            return datetime.strptime(s.split(" ")[0], "%d/%m/%Y").strftime("%d/%m/%Y")
        if len(s) >= 10:
            return s[:10]
        return s
    except:
        return val

date_cols = []
for col in df_out.columns:
    if "log in date" in col.lower() or "etd" in col.lower():
        date_cols.append(col)

for col in date_cols:
    df_out[col] = df_out[col].apply(only_date)

# ==== Chế độ cập nhật giữ lại dòng đã có status ====
if os.path.exists(excel_file_out):
    df_exist = pd.read_excel(excel_file_out)
    # Tìm cột report/status trong file cũ
    report_col_exist = None
    status_col_exist = None
    for col in df_exist.columns:
        if "report" in col.lower():
            report_col_exist = col
        if "status" in col.lower():
            status_col_exist = col
    # Set các mã đã có status (không trắng)
    reports_with_status = set(
        df_exist[
            df_exist[status_col_exist].notnull() &
            (df_exist[status_col_exist].astype(str).str.strip() != "")
        ][report_col_exist].astype(str)
    )
    # Chỉ lấy report chưa có status để update/copy mới
    df_out_new = df_out[~df_out[report_col_main].astype(str).isin(reports_with_status)]
    # Giữ lại các dòng cũ đã có status
    df_keep = df_exist[df_exist[report_col_exist].astype(str).isin(reports_with_status)].copy()
    # Đảm bảo cột đủ đầy
    for col in df_out_new.columns:
        if col not in df_keep.columns:
            df_keep[col] = ""
    for col in df_keep.columns:
        if col not in df_out_new.columns:
            df_out_new[col] = ""
    df_out_new = df_out_new[df_keep.columns]
    df_final = pd.concat([df_keep, df_out_new], ignore_index=True)
else:
    df_final = df_out

# ==== Xuất file CHÍNH ====
df_final.to_excel(excel_file_out, index=False)
wb = load_workbook(excel_file_out)
ws = wb.active

# ==== Format Excel ====
thin = Side(border_style="thin", color="888888")
fill_late = PatternFill("solid", fgColor="FF6961")
fill_due = PatternFill("solid", fgColor="FFEB7A")
fill_must = PatternFill("solid", fgColor="FFA54F")
fill_complete = PatternFill("solid", fgColor="D3D3D3")
header_fill = PatternFill("solid", fgColor="B7E1CD")

header = [cell.value for cell in ws[1]]
status_col_idx = None
for i, name in enumerate(header):
    if name and str(name).strip().lower() == status_col.strip().lower():
        status_col_idx = i + 1
        break

valid_statuses = ("active", "pending", "late", "due", "must")
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter
    for cell in col:
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        if cell.row == 1:
            cell.font = Font(bold=True)
            cell.fill = header_fill
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max(max_length + 2, 15)

for row in range(2, ws.max_row + 1):
    ws.row_dimensions[row].height = 90
    status_val = ws.cell(row=row, column=status_col_idx).value
    status_val_norm = str(status_val).strip().lower() if status_val else ""
    fill = None
    if status_val_norm == "late":
        fill = fill_late
    elif status_val_norm == "due":
        fill = fill_due
    elif status_val_norm == "must":
        fill = fill_must
    elif status_val_norm == "complete":
        fill = fill_complete
    if status_val_norm not in valid_statuses:
        ws.row_dimensions[row].hidden = True
    if fill:
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

wb.save(excel_file_out)
print(f"✅ ĐÃ XUẤT FILE: {excel_file_out}")
