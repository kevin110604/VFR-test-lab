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
import numpy as np
import re

# ========= Helpers cho phần NGÀY =========
CURRENT_YEAR = datetime.now().year

def normalize_col(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())

def find_login_date_col(df: pd.DataFrame):
    # Ưu tiên các biến thể phổ biến
    targets = {
        "log in date", "login date", "log-in date", "logindate", "log in-date",
        "logged in date", "log_date", "log date"
    }
    cmap = {c: normalize_col(c) for c in df.columns}
    for orig, low in cmap.items():
        if low in targets:
            return orig
    # fallback: chứa cả 'log' và 'date'
    for orig, low in cmap.items():
        if "log" in low and "date" in low:
            return orig
    return None

_has_year_pat = re.compile(r"\b(\d{4}|\d{2})\b")
_missing_year_pat = re.compile(r"^\s*(\d{1,2})[./\- ]([A-Za-z]{3}|\d{1,2})\s*$")  # dd-MMM, dd/MM, dd.MM, dd-MM

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct":10, "nov":11, "dec":12
}

def _month_to_int(token: str):
    t = token.strip().lower()
    if t.isdigit():
        m = int(t)
        return m if 1 <= m <= 12 else None
    return MONTH_MAP.get(t[:3])  # "aug", "dec" ...

def _attach_current_year_if_missing(text: str) -> str:
    """
    Gắn năm theo quy tắc:
      - Không có năm:
          * Tháng 1  => năm hiện tại
          * Tháng 12 => năm hiện tại - 1
          * Tháng khác => năm hiện tại
      - Giữ nguyên nếu đã có năm.
    """
    s = (text or "").strip()
    if not s:
        return s
    # đã có năm
    if _has_year_pat.search(s):
        return s

    # cố gắng bắt dd-sep-mm(token) với tháng dạng chữ hoặc số
    m = _missing_year_pat.match(s)
    if m:
        day = m.group(1)
        mon_token = m.group(2)
        mon = _month_to_int(mon_token)
        if mon is None:
            # không nhận ra tháng => gắn năm hiện tại ở cuối
            return f"{s} {CURRENT_YEAR}"

        if mon == 12:
            year = CURRENT_YEAR - 1
        elif mon == 1:
            year = CURRENT_YEAR
        else:
            year = CURRENT_YEAR

        # nếu tháng dạng chữ thì dùng dd-MMM-YYYY, nếu số thì dd/MM/YYYY
        if mon_token.isdigit():
            s2 = re.sub(r"[.\- ]", "/", s)  # chuẩn hoá phân cách
            return f"{s2}/{year}"
        else:
            # dùng '-' để nhất quán khi có tháng chữ
            return f"{s}-{year}"

    # không khớp pattern, vẫn gắn năm hiện tại
    return f"{s} {CURRENT_YEAR}"

def parse_login_dates(series: pd.Series) -> pd.Series:
    # Thêm năm theo quy tắc Jan/Dec rồi parse
    tmp = series.astype(str).map(_attach_current_year_if_missing)
    dt = pd.to_datetime(tmp, errors="coerce", dayfirst=True, infer_datetime_format=True)
    if dt.isna().mean() > 0.8:
        dt2 = pd.to_datetime(tmp, errors="coerce", dayfirst=False, infer_datetime_format=True)
        if dt2.notna().sum() > dt.notna().sum():
            dt = dt2
    # NaT coi là rất cũ để bị loại khi chỉ giữ phần cuối
    return dt.fillna(pd.Timestamp(1900, 1, 1))

# ==== HÀM ĐẢM BẢO FOLDER SHAREPOINT TỒN TẠI (TỰ TẠO TỪNG CẤP) ====
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

# ==== Cấu hình SharePoint ====
site_url = "https://jonathancharles.sharepoint.com/sites/TESTLAB-VFR9"
username = "tan_qa@vfr.net.vn"
password = "qaz@Tat@123"

# ================= PHẦN 1: XUẤT FILE DS SẢN PHẨM TEST VỚI QR =================
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

# Lọc dòng: loại bỏ outsource
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

# Xác định cột report là cột số 2 trong file ds
report_col_main = df_out.columns[1]

# Chuẩn hoá và xoá trùng mã report, giữ dòng cuối
df_out[report_col_main] = df_out[report_col_main].astype(str).str.strip()
df_out = df_out.drop_duplicates(subset=report_col_main, keep='last')

# Lọc theo Report ID >= 4500
df_out["__report_num__"] = pd.to_numeric(df_out[report_col_main].astype(str).str.extract(r'(\d+)$')[0], errors="coerce")
df_out = df_out[df_out["__report_num__"] >= 4500]
df_out.drop(columns="__report_num__", inplace=True)

# Thêm cột bổ sung (chỉ thêm Test Date, Complete Date)
for col in ["Test Date", "Complete Date"]:
    df_out[col] = ""

# Tạo cột QR Code link
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

# ============= PHẦN SỬA ĐỂ KHÔNG GIỮ LẠI DÒNG RATING BLANK =============
def is_valid_rating(val):
    if pd.isnull(val): return False
    sval = str(val).strip()
    if sval == "": return False
    if sval.lower() == "nan": return False
    return True

if os.path.exists(excel_file_out):
    df_exist = pd.read_excel(excel_file_out)
    # Chuẩn hoá tên cột ở cả hai DataFrame
    df_exist.columns = [col.strip() for col in df_exist.columns]
    df_out.columns = [col.strip() for col in df_out.columns]

    # Xác định cột mã report và cột rating (trong file cũ)
    report_col_exist = None
    rating_col_exist = None
    for col in df_exist.columns:
        if "report" in col.lower() and ("#" in col or "id" in col or "report" == col.lower().strip()):
            report_col_exist = col
        if "rating" in col.lower():
            rating_col_exist = col

    # Dùng đúng report_col_main là cột số 2
    report_col_main = df_out.columns[1]

    # Chuẩn hoá giá trị index (report)
    df_exist[report_col_exist] = df_exist[report_col_exist].astype(str).str.strip()
    df_out[report_col_main] = df_out[report_col_main].astype(str).str.strip()
    # Xoá trùng mã report trên df_out (đề phòng file cũ)
    df_out = df_out.drop_duplicates(subset=report_col_main, keep='last')

    # Build mapping từ mã report -> dữ liệu SharePoint
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
                    # Chỉ update nếu cột này có trong dữ liệu mới
                    if col != rating_col_exist and col in new_data:
                        row[col] = new_data[col]
            updated_rows.append(row)
    df_final = pd.DataFrame(updated_rows, columns=df_exist.columns)
else:
    df_final = df_out

df_final.to_excel(excel_file_out, index=False)

# ============== CÁC ĐOẠN ĐỊNH DẠNG, TÔ MÀU, UPLOAD SHAREPOINT GIỮ NGUYÊN ==============
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
            for c in range(1, ws.max_column + 1):
                ws.cell(row=row, column=c).fill = fill

wb.save(excel_file_out)
print("Da xuat file: {}".format(excel_file_out))

# ================= PHẦN 2: XỬ LÝ FILE completed_items.xlsx LOCAL =================
completed_file = "completed_items.xlsx"
if not os.path.exists(completed_file):
    print("Khong tim thay file {} o local. Khong up len SharePoint.".format(completed_file))
else:
    df = pd.read_excel(completed_file, dtype=str)  # Đọc tất cả dạng chuỗi

    keep_n = 200
    login_col = find_login_date_col(df)
    if login_col is None:
        # Không tìm thấy cột log in date => fallback giữ 200 dòng cuối theo vị trí
        total_rows = len(df)
        if total_rows > keep_n:
            df = df.iloc[-keep_n:].copy()
            print(f"[completed_items] Khong tim thay cot 'log in date'. Da xoa {total_rows - keep_n} dong dau, giu {keep_n} dong cuoi.")
        else:
            print(f"[completed_items] File hien chi co {total_rows} dong, khong can xoa.")
    else:
        # Sắp xếp theo log in date (thêm năm theo quy tắc Jan/Dec)
        dt = parse_login_dates(df[login_col])
        df = df.assign(__login_dt__=dt).sort_values(["__login_dt__", login_col], ascending=[True, True], kind="mergesort")
        total_rows = len(df)
        if total_rows > keep_n:
            df = df.tail(keep_n).copy()
            print(f"[completed_items] Da xoa {total_rows - keep_n} dong cu (theo '{login_col}'), giu {keep_n} dong moi nhat.")
        else:
            print(f"[completed_items] File hien chi co {total_rows} dong, khong can xoa.")
        df.drop(columns="__login_dt__", inplace=True, errors="ignore")

    # Xuất lại file giữ nguyên heading và định dạng
    df.to_excel(completed_file, index=False)

    # Định dạng lại file
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

# ================= PHẦN 3: XỬ LÝ FILE TRF.xlsx LOCAL =================
trf_file = "TRF.xlsx"
if not os.path.exists(trf_file):
    print("Khong tim thay file {} o local. Khong up len SharePoint.".format(trf_file))
else:
    df = pd.read_excel(trf_file, dtype=str)

    keep_n = 200
    login_col = find_login_date_col(df)
    if login_col is None:
        total_rows = len(df)
        if total_rows > keep_n:
            df = df.iloc[-keep_n:].copy()
            print(f"[TRF] Khong tim thay cot 'log in date'. Da xoa {total_rows - keep_n} dong dau, giu {keep_n} dong cuoi.")
        else:
            print(f"[TRF] File hien chi co {total_rows} dong, khong can xoa.")
    else:
        dt = parse_login_dates(df[login_col])
        df = df.assign(__login_dt__=dt).sort_values(["__login_dt__", login_col], ascending=[True, True], kind="mergesort")
        total_rows = len(df)
        if total_rows > keep_n:
            df = df.tail(keep_n).copy()
            print(f"[TRF] Da xoa {total_rows - keep_n} dong cu (theo '{login_col}'), giu {keep_n} dong moi nhat.")
        else:
            print(f"[TRF] File hien chi co {total_rows} dong, khong can xoa.")
        df.drop(columns="__login_dt__", inplace=True, errors="ignore")

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
