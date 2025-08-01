import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import io
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from datetime import datetime, timedelta
import numpy as np

# ==== HAM DAM BAO FOLDER SHAREPOINT TON TAI (TU TAO TUNG CAP) ====
def ensure_folder(ctx, folder_url):
    folder_url = folder_url.rstrip("/")
    root_url = "/".join(folder_url.strip("/").split("/")[:4])  # /sites/...
    parts = folder_url.strip("/").split("/")[4:]  # sau /sites/...
    current_url = root_url
    for part in parts:
        current_url = current_url + "/" + part
        try:
            ctx.web.folders.add(current_url).execute_query()
        except Exception as e:
            if "already exists" not in str(e).lower():
                print("Loi tao folder {}: {}".format(current_url, e))
                raise
    return ctx.web.get_folder_by_server_relative_url(folder_url)

# ==== Cau hinh SharePoint ====
site_url = "https://jonathancharles.sharepoint.com/sites/TESTLAB-VFR9"
username = "tan_qa@vfr.net.vn"
password = "qaz@Tat@123"

# ================= PHAN 1: XUAT FILE DS SAN PHAM TEST VOI QR =================
relative_url = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/QAD-Outstanding list-2025.xlsx"
excel_file_out = "ds san pham test voi qr.xlsx"

ctx_auth = AuthenticationContext(site_url)
if not ctx_auth.acquire_token_for_user(username, password):
    raise Exception("Khong ket noi duoc SharePoint!")
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
    raise Exception("Khong tim thay sheet 'OL list'!")
df = excel_file.parse(sheet_name)

cols_selected = list(df.columns[1:23])
rating_col = None
for col in df.columns:
    if "rating" in col.lower():
        rating_col = col
        break
if rating_col and rating_col not in cols_selected:
    cols_selected.append(rating_col)

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

# Loc dong: loai bo outsource
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

# Xac dinh cot report la cot so 2 trong file ds
report_col_main = df_out.columns[1]

# Chuan hoa va xoa trung ma report, giu dong cuoi
df_out[report_col_main] = df_out[report_col_main].astype(str).str.strip()
df_out = df_out.drop_duplicates(subset=report_col_main, keep='last')

# Loc theo Report ID >= 4500
df_out["__report_num__"] = pd.to_numeric(df_out[report_col_main].astype(str).str.extract(r'(\d+)$')[0], errors="coerce")
df_out = df_out[df_out["__report_num__"] >= 4500]
df_out.drop(columns="__report_num__", inplace=True)

# Them cot bo sung (chi them Test Date, Complete Date)
for col in ["Test Date", "Complete Date"]:
    df_out[col] = ""

# Tao cot QR Code link
qr_url_dict = {}
for idx, row in df_out.iterrows():
    report_raw = str(row[report_col_main]).strip()
    if report_raw and report_raw.lower() != "nan":
        url = f"http://103.77.166.187:2004/update?report={report_raw}"
        qr_url_dict[report_raw] = url
df_out["QR Code"] = df_out[report_col_main].astype(str).map(qr_url_dict).fillna("")

def only_date(val):
    if pd.isnull(val) or not str(val).strip():
        return ""
    try:
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
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

### ============= PHAN SUA DE KHONG GIU LAI DONG RATING BLANK =============
def is_valid_rating(val):
    if pd.isnull(val): return False
    sval = str(val).strip()
    if sval == "": return False
    if sval.lower() == "nan": return False
    return True

if os.path.exists(excel_file_out):
    df_exist = pd.read_excel(excel_file_out)
    # Chuan hoa ten cot o ca hai DataFrame (nen lam)
    df_exist.columns = [col.strip() for col in df_exist.columns]
    df_out.columns = [col.strip() for col in df_out.columns]

    # Xac dinh cot ma report va cot rating (trong file cu)
    report_col_exist = None
    rating_col_exist = None
    for col in df_exist.columns:
        if "report" in col.lower() and ("#" in col or "id" in col or "report" == col.lower().strip()):
            report_col_exist = col
        if "rating" in col.lower():
            rating_col_exist = col

    # Dung dung report_col_main la cot so 2
    report_col_main = df_out.columns[1]

    # Chuan hoa gia tri index (report)
    df_exist[report_col_exist] = df_exist[report_col_exist].astype(str).str.strip()
    df_out[report_col_main] = df_out[report_col_main].astype(str).str.strip()
    # Xoa trung ma report tren df_out (de phong file cu)
    df_out = df_out.drop_duplicates(subset=report_col_main, keep='last')

    # Build mapping tu ma report -> du lieu SharePoint
    sharepoint_data = df_out.set_index(report_col_main).to_dict(orient="index")

    updated_rows = []
    for idx, row in df_exist.iterrows():
        report_val = row[report_col_exist]
        rating_val = row[rating_col_exist]
        if is_valid_rating(rating_val):
            updated_rows.append(row)
        else:
            report_str = str(report_val).strip()
            if pd.notnull(report_val) and report_str in sharepoint_data:
                new_data = sharepoint_data[report_str]
                for col in df_exist.columns:
                    # Chi update neu cot nay co trong du lieu moi
                    if col != rating_col_exist and col in new_data:
                        row[col] = new_data[col]
            updated_rows.append(row)
    df_final = pd.DataFrame(updated_rows, columns=df_exist.columns)
else:
    df_final = df_out

df_final.to_excel(excel_file_out, index=False)

# ============== CAC DOAN DINH DANG, TO MAU, UPLOAD SHAREPOINT GIU NGUYEN ==============
wb = load_workbook(excel_file_out)
ws = wb.active

thin = Side(border_style="thin", color="888888")
fill_late = PatternFill("solid", fgColor="FF6961")
fill_due = PatternFill("solid", fgColor="FFEB7A")
fill_must = PatternFill("solid", fgColor="FFA54F")
fill_complete = PatternFill("solid", fgColor="D3D3D3")
header_fill = PatternFill("solid", fgColor="B7E1CD")

header = [cell.value for cell in ws[1]]
status_col_idx = None
for i, name in enumerate(header):
    if name and str(name).strip().lower() == (status_col.strip().lower() if status_col else ""):
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
    if status_col_idx:
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
print("Da xuat file: {}".format(excel_file_out))

# ================= PHAN 2: XU LY FILE completed_items.xlsx LOCAL =================
completed_file = "completed_items.xlsx"
if not os.path.exists(completed_file):
    print("Khong tim thay file {} o local. Khong up len SharePoint.".format(completed_file))
else:
    df = pd.read_excel(completed_file, dtype=str)  # Doc tat ca duoi dang chuoi

    max_rows = 200
    total_rows = len(df)

    if total_rows > max_rows:
        # Giu 200 dong cuoi cung (moi nhat)
        df = df.iloc[-max_rows:].copy()
        print("Da xoa {} dong cu, chi giu lai {} ma report moi nhat trong file {}.".format(
            total_rows - max_rows, max_rows, completed_file))
    else:
        print("File hien chi co {} dong, khong can xoa dong nao.".format(total_rows))

    # Xuat lai file giu nguyen heading va dinh dang
    df.to_excel(completed_file, index=False)

    # Dinh dang lai file (neu muon)
    wb = load_workbook(completed_file)
    ws = wb.active

    thin = Side(border_style="thin", color="888888")
    header_fill = PatternFill("solid", fgColor="B7E1CD")

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

    wb.save(completed_file)
    print("Da xuat file: {}".format(completed_file))
    
    # ==== UPLOAD FILE LEN SHAREPOINT ====
    upload_relative_url = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/completed_items.xlsx"
    folder_excel = os.path.dirname(upload_relative_url)
    ensure_folder(ctx, folder_excel)
    if os.path.exists(completed_file):
        with open(completed_file, "rb") as f:
            ctx.web.get_folder_by_server_relative_url(folder_excel) \
                .upload_file(os.path.basename(upload_relative_url), f.read()).execute_query()
        print("Da upload file {} len SharePoint: {}".format(completed_file, upload_relative_url))
    else:
        print("File completed_items.xlsx chua ton tai, khong the upload.")

trf_file = "TRF.xlsx"
if not os.path.exists(trf_file):
    print("Khong tim thay file {} o local. Khong up len SharePoint.".format(trf_file))
else:
    df = pd.read_excel(trf_file, dtype=str)

    max_rows = 200
    total_rows = len(df)

    if total_rows > max_rows:
        df = df.iloc[-max_rows:].copy()
        print("Da xoa {} dong cu, chi giu lai {} ma report moi nhat trong file {}.".format(
            total_rows - max_rows, max_rows, trf_file))
    else:
        print("File hien chi co {} dong, khong can xoa dong nao.".format(total_rows))

    df.to_excel(trf_file, index=False)

    wb = load_workbook(trf_file)
    ws = wb.active

    thin = Side(border_style="thin", color="888888")
    header_fill = PatternFill("solid", fgColor="B7E1CD")

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

    wb.save(trf_file)
    print("Da xuat file: {}".format(trf_file))
    
    # ==== UPLOAD FILE LEN SHAREPOINT ====
    upload_relative_url_trf = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/TRF.xlsx"
    folder_excel_trf = os.path.dirname(upload_relative_url_trf)
    ensure_folder(ctx, folder_excel_trf)
    if os.path.exists(trf_file):
        with open(trf_file, "rb") as f:
            ctx.web.get_folder_by_server_relative_url(folder_excel_trf) \
                .upload_file(os.path.basename(upload_relative_url_trf), f.read()).execute_query()
        print("Da upload file {} len SharePoint: {}".format(trf_file, upload_relative_url_trf))
    else:
        print("File TRF.xlsx chua ton tai, khong the upload.")
