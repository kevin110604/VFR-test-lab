import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import os
import io
import json
import re
from datetime import datetime, timedelta, date
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# SharePoint CSOM
from office365.sharepoint.client_context import ClientContext

# MSAL cho OAuth (Delegated + cache)
import msal

# ================== CẤU HÌNH OAUTH (DELEGATED + CACHE) ==================
TENANT_ID = "064944f6-1e04-4050-b3e1-e361758625ec"       # Directory (tenant) ID
CLIENT_ID = "9abf6ee2-50c8-47c8-a9f2-8cf18587c6ea"       # Application (client) ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SP_HOST = "https://jonathancharles.sharepoint.com"

# Scope hợp lệ của SharePoint (Delegated) - KHÔNG dùng openid/profile/offline_access ở đây
SPO_SCOPES = [
    f"{SP_HOST}/AllSites.Read",
    f"{SP_HOST}/AllSites.Write",
]

TOKEN_CACHE_FILE = "token_cache.bin"


def _load_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def acquire_spo_access_token() -> str:
    """
    Lấy access token cho SharePoint Online (Delegated).
    - Ưu tiên: acquire_token_silent() từ cache
    - Nếu chưa có/expired: Device Code Flow (in mã ra console), login 1 lần
    - Cache tự lưu để lần sau silent
    """
    cache = _load_token_cache()
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SPO_SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SPO_SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Không khởi tạo được device flow: {json.dumps(flow, indent=2)}")
        # Hướng dẫn login 1 lần (copy link + code ra trình duyệt)
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    _save_token_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(f"Không lấy được access token: {result.get('error_description', str(result))}")

    return result["access_token"]


def get_ctx(site_url: str) -> ClientContext:
    """
    ClientContext sử dụng custom authenticate_request:
    - Mỗi request tự gắn 'Authorization: Bearer <token>' lấy từ MSAL cache
    - Tránh hoàn toàn bug timezone (naive/aware) trong lib
    """
    ctx = ClientContext(site_url)

    def _auth(request):
        token = acquire_spo_access_token()  # lấy từ cache, tự refresh khi cần
        request.ensure_header("Authorization", "Bearer " + token)

    # Ghi đè cơ chế auth mặc định
    ctx.authentication_context.authenticate_request = _auth
    print("AUTH MODE:", "Delegated (Device Code + Token Cache) [custom auth]")
    return ctx


# ================== PHẦN SHAREPOINT / XỬ LÝ EXCEL ==================
CURRENT_YEAR = datetime.now().year

def normalize_col(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())

def find_login_date_col(df: pd.DataFrame):
    targets = {
        "log in date", "login date", "log-in date", "logindate", "log in-date",
        "logged in date", "log_date", "log date"
    }
    cmap = {c: normalize_col(c) for c in df.columns}
    for orig, low in cmap.items():
        if low in targets:
            return orig
    for orig, low in cmap.items():
        if "log" in low and "date" in low:
            return orig
    return None

_has_year_pat = re.compile(r"\b(\d{4}|\d{2})\b")
_missing_year_pat = re.compile(r"^\s*(\d{1,2})[./\- ]([A-Za-z]{3}|\d{1,2})\s*$")
MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct":10, "nov":11, "dec":12
}

def _month_to_int(token: str):
    t = token.strip().lower()
    if t.isdigit():
        m = int(t)
        return m if 1 <= m <= 12 else None
    return MONTH_MAP.get(t[:3])

def _attach_current_year_if_missing(text: str) -> str:
    s = (text or "").strip()
    if not s:
        return s
    if _has_year_pat.search(s):
        return s
    m = _missing_year_pat.match(s)
    if m:
        mon_token = m.group(2)
        mon = _month_to_int(mon_token)
        year = CURRENT_YEAR if mon != 12 else CURRENT_YEAR - 1
        if mon_token.isdigit():
            s2 = re.sub(r"[.\- ]", "/", s)
            return f"{s2}/{year}"
        else:
            return f"{s}-{year}"
    return f"{s} {CURRENT_YEAR}"

def parse_login_dates(series: pd.Series) -> pd.Series:
    """
    Ưu tiên parse ISO '%Y-%m-%d %H:%M:%S' (không cảnh báo),
    nếu không được thì fallback dayfirst True/False.
    """
    tmp = series.astype(str).map(_attach_current_year_if_missing)

    # Ưu tiên parse ISO chuẩn
    dt = pd.to_datetime(tmp, errors="coerce", format="%Y-%m-%d %H:%M:%S", exact=False)
    if dt.isna().any():
        # Thử parse generic với dayfirst
        dt2 = pd.to_datetime(tmp, errors="coerce", dayfirst=True, infer_datetime_format=True)
        if dt2.notna().sum() > dt.notna().sum():
            dt = dt2
        else:
            # fallback cuối cùng
            dt3 = pd.to_datetime(tmp, errors="coerce", dayfirst=False, infer_datetime_format=True)
            if dt3.notna().sum() > dt.notna().sum():
                dt = dt3
    return dt.fillna(pd.Timestamp(1900, 1, 1))

def hide_rows_by_login_date(ws, login_col_name_or_idx, today=None):
    """
    Ẩn các dòng có ngày đăng nhập (login date) <= today - 2 ngày.
    Parse ngày theo thứ tự: ISO -> dayfirst=True -> dayfirst=False.
    """
    if today is None:
        today = datetime.now().date()
    hide_threshold = today - timedelta(days=2)
    if isinstance(login_col_name_or_idx, int):
        login_col_idx = login_col_name_or_idx
    else:
        login_col_idx = None
        header = [cell.value for cell in ws[1]]
        for i, name in enumerate(header, start=1):
            if name and normalize_col(name) == normalize_col(login_col_name_or_idx):
                login_col_idx = i
                break
        if login_col_idx is None:
            for i, name in enumerate(header, start=1):
                if name and ("log" in normalize_col(name) and "date" in normalize_col(name)):
                    login_col_idx = i
                    break
    if not login_col_idx:
        return
    for r in range(2, ws.max_row + 1):
        raw = ws.cell(row=r, column=login_col_idx).value
        raw_str = "" if raw is None else str(raw).strip()
        if not raw_str:
            continue
        s_with_year = _attach_current_year_if_missing(raw_str)

        # Try ISO first
        try:
            dt = pd.to_datetime(s_with_year, errors="coerce", format="%Y-%m-%d %H:%M:%S")
        except Exception:
            dt = None
        if dt is None or pd.isna(dt):
            dt = pd.to_datetime(s_with_year, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            dt = pd.to_datetime(s_with_year, errors="coerce", dayfirst=False)
        if pd.isna(dt):
            continue
        if dt.date() <= hide_threshold:
            ws.row_dimensions[r].hidden = True

def ensure_folder(ctx, folder_url):
    folder_url = folder_url.rstrip("/")
    root_url = "/".join(folder_url.strip("/").split("/")[:4])
    parts = folder_url.strip("/").split("/")[4:]
    current_url = root_url
    for part in parts:
        current_url = current_url + "/" + part
        try:
            ctx.web.folders.add(current_url).execute_query()
        except Exception as e:
            if "already exists" not in str(e).lower() and "conflict" not in str(e).lower():
                print(f"Loi tao folder {current_url}: {e}")
                raise
    return ctx.web.get_folder_by_server_relative_url(folder_url)

# ====== HỖ TRỢ: Định dạng cột ngày thành dd-mmm cho openpyxl ======
def _parse_to_datetime_or_none(text: str):
    """
    Cố gắng parse chuỗi sang datetime (không timezone).
    Ưu tiên ISO '%Y-%m-%d %H:%M:%S', sau đó thử dayfirst True/False.
    """
    s = (text or "").strip()
    if not s:
        return None
    # ISO trước
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    # Pandas fallback
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if pd.isna(dt):
        return None
    return pd.Timestamp(dt).to_pydatetime()

def format_date_columns(ws, target_headers=("log in date", "etd"), number_format="DD-MMM"):
    """
    - Tìm các cột header nằm trong target_headers (so sánh normalize).
    - Với từng ô dữ liệu:
        + Nếu là datetime/date: set number_format = 'DD-MMM'
        + Nếu là chuỗi: thử parse; nếu parse được -> gán datetime + number_format
        + Nếu đã đúng định dạng thì set lại number_format (idempotent, an toàn)
    """
    header_cells = [cell.value for cell in ws[1]]
    target_norm = {normalize_col(h) for h in target_headers}
    target_col_indices = []
    for idx, name in enumerate(header_cells, start=1):
        if name and normalize_col(name) in target_norm:
            target_col_indices.append(idx)

    for col_idx in target_col_indices:
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=col_idx)
            val = cell.value
            if val is None or (isinstance(val, str) and not val.strip()):
                continue
            if isinstance(val, datetime):
                # đã là datetime -> chỉ cần set number_format
                cell.number_format = number_format
                continue
            if isinstance(val, date):
                # date -> cast sang datetime để Excel hiển thị đúng
                cell.value = datetime(val.year, val.month, val.day)
                cell.number_format = number_format
                continue
            if isinstance(val, (int, float)):
                # có thể là serial date; để nguyên, chỉ set number_format
                cell.number_format = number_format
                continue
            # Nếu là chuỗi -> cố parse
            if isinstance(val, str):
                dt = _parse_to_datetime_or_none(val)
                if dt is not None:
                    cell.value = dt
                    cell.number_format = number_format
                # nếu không parse được thì bỏ qua (giữ nguyên)

# ==== Cấu hình SharePoint ====
site_url = f"{SP_HOST}/sites/TESTLAB-VFR9"

# ================= PHẦN 1: XUẤT FILE DS SẢN PHẨM TEST VỚI QR =================
relative_url = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/QAD-Outstanding list-2025.xlsx"
excel_file_out = "ds san pham test voi qr.xlsx"

# --- KẾT NỐI SharePoint bằng Delegated (Device Code + Cache) ---
ctx = get_ctx(site_url)

download = io.BytesIO()
ctx.web.get_file_by_server_relative_url(relative_url).download(download).execute_query()
download.seek(0)

excel_file = pd.ExcelFile(download)
sheet_name = next((n for n in excel_file.sheet_names if n.strip().lower() == "ol list"), None)
if not sheet_name:
    raise Exception("Khong tim thay sheet 'OL list'!")
df = excel_file.parse(sheet_name)

cols_selected = list(df.columns[1:23])
rating_col = next((c for c in df.columns if "rating" in str(c).lower()), None)
if rating_col and rating_col not in cols_selected:
    cols_selected.append(rating_col)

def find_col(keywords):
    for col in df.columns:
        if all(k in str(col).lower() for k in keywords):
            return col
    for col in df.columns:
        if any(k in str(col).lower() for k in keywords):
            return col
    return None

type_of_cols = [c for c in df.columns if "type of" in str(c).lower()]
status_col = find_col(["status"])
report_col = find_col(["report", "#"])

def norm(s):
    return (str(s).strip().lower() if pd.notnull(s) else "")

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
df_out = df.iloc[rows].copy()[cols_selected]

report_col_main = df_out.columns[1]
df_out[report_col_main] = df_out[report_col_main].astype(str).str.strip()
df_out = df_out.drop_duplicates(subset=report_col_main, keep='last')

df_out["__report_num__"] = pd.to_numeric(df_out[report_col_main].astype(str).str.extract(r'(\d+)$')[0], errors="coerce")
df_out = df_out[df_out["__report_num__"] >= 4500].drop(columns="__report_num__")

for col in ["Test Date", "Complete Date"]:
    df_out[col] = ""

qr_url_dict = {}
for _, row in df_out.iterrows():
    report_raw = str(row[report_col_main]).strip()
    if report_raw and report_raw.lower() != "nan":
        qr_url_dict[report_raw] = f"http://103.77.166.187:8246/update?report={report_raw}"
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
            return datetime.strptime(s.split(" ")[0], "%d/%m/%Y")
        if len(s) >= 10:
            return s[:10]
        return s
    except:
        return val

date_cols = [c for c in df_out.columns if "log in date" in str(c).lower() or "etd" in str(c).lower()]
for c in date_cols:
    df_out[c] = df_out[c].apply(only_date)

def is_valid_rating(val):
    if pd.isnull(val): return False
    sval = str(val).strip().lower()
    return sval != "" and sval != "nan"

if os.path.exists(excel_file_out):
    df_exist = pd.read_excel(excel_file_out)
    df_exist.columns = [c.strip() for c in df_exist.columns]
    df_out.columns = [c.strip() for c in df_out.columns]

    report_col_exist = next((c for c in df_exist.columns if "report" in c.lower() and ("#" in c or "id" in c or c.lower().strip()=="report")), None)
    rating_col_exist = next((c for c in df_exist.columns if "rating" in c.lower()), None)

    report_col_main = df_out.columns[1]
    df_exist[report_col_exist] = df_exist[report_col_exist].astype(str).str.strip()
    df_out[report_col_main] = df_out[report_col_main].astype(str).str.strip()
    df_out = df_out.drop_duplicates(subset=report_col_main, keep='last')

    sharepoint_data = df_out.set_index(report_col_main).to_dict(orient="index")

    updated = []
    for _, row in df_exist.iterrows():
        report_val = row[report_col_exist]
        rating_val = row[rating_col_exist]
        if is_valid_rating(rating_val):
            updated.append(row)
        else:
            key = str(report_val).strip()
            if pd.notnull(report_val) and key in sharepoint_data:
                new_data = sharepoint_data[key]
                for col in df_exist.columns:
                    if col != rating_col_exist and col in new_data:
                        row[col] = new_data[col]
            updated.append(row)
    df_final = pd.DataFrame(updated, columns=df_exist.columns)
else:
    df_final = df_out

df_final.to_excel(excel_file_out, index=False)

# ======= ĐỊNH DẠNG, TÔ MÀU =======
wb = load_workbook(excel_file_out)
ws = wb.active

thin = Side(border_style="thin", color="888888")
fill_late = PatternFill("solid", fgColor="FF6961")
fill_due = PatternFill("solid", fgColor="FFEB7A")
fill_must = PatternFill("solid", fgColor="FFA54F")
fill_complete = PatternFill("solid", fgColor="D3D3D3")
header_fill = PatternFill("solid", fgColor="B7E1CD")

header = [cell.value for cell in ws[1]]
# status_col có thể None nếu không tìm thấy
status_col_idx = None
# Nếu bạn muốn cố định tên cột, có thể set status_col = "Status"
# Ở đây mình giữ nguyên: nếu có biến status_col ở find_col phía trên thì dùng
try:
    if 'status_col' in locals() and status_col:
        for i, name in enumerate(header):
            if name and str(name).strip().lower() == status_col.strip().lower():
                status_col_idx = i + 1
                break
except Exception:
    status_col_idx = None

valid_statuses = ("active", "pending", "late", "due", "must", "complete")
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
        status_val_norm = str(ws.cell(row=row, column=status_col_idx).value or "").strip().lower()
        fill = None
        if status_val_norm == "late":
            fill = fill_late
        elif status_val_norm == "due":
            fill = fill_due
        elif status_val_norm == "must":
            fill = fill_must
        elif status_val_norm == "complete":
            fill = fill_complete
        if status_val_norm and status_val_norm not in valid_statuses:
            ws.row_dimensions[row].hidden = True
        if fill:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=row, column=c).fill = fill

wb.save(excel_file_out)
print("Da xuat file:", excel_file_out)

# ================= PHẦN 2: completed_items.xlsx =================
completed_file = "completed_items.xlsx"
if not os.path.exists(completed_file):
    print(f"Khong tim thay file {completed_file} o local. Khong up len SharePoint.")
else:
    df_cpl = pd.read_excel(completed_file, dtype=str)
    login_col_cpl = find_login_date_col(df_cpl)
    df_cpl.to_excel(completed_file, index=False)

    wb = load_workbook(completed_file); ws = wb.active
    header_fill = PatternFill("solid", fgColor="B7E1CD")

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            if cell.row == 1:
                cell.font = Font(bold=True); cell.fill = header_fill
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_length + 2, 15)

    if login_col_cpl:
        hide_rows_by_login_date(ws, login_col_cpl, today=datetime.now().date())

    wb.save(completed_file)
    print("Da xuat file:", completed_file)

    upload_relative_url = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/completed_items.xlsx"
    folder_excel = os.path.dirname(upload_relative_url)
    ensure_folder(ctx, folder_excel)
    with open(completed_file, "rb") as f:
        ctx.web.get_folder_by_server_relative_url(folder_excel)\
           .upload_file(os.path.basename(upload_relative_url), f.read()).execute_query()
    print("Da upload file", completed_file, "len SharePoint:", upload_relative_url)

# ================= PHẦN 3: TRF.xlsx =================
trf_file = "TRF.xlsx"
if not os.path.exists(trf_file):
    print(f"Khong tim thay file {trf_file} o local. Khong up len SharePoint.")
else:
    # 1) Đọc và ghi lại để giữ cấu trúc/columns
    df_trf = pd.read_excel(trf_file, dtype=str)
    login_col_trf = find_login_date_col(df_trf)
    df_trf.to_excel(trf_file, index=False)

    # 2) Mở bằng openpyxl để định dạng header + căn giữa + khung, rồi format ngày
    wb = load_workbook(trf_file); ws = wb.active
    header_fill = PatternFill("solid", fgColor="B7E1CD")

    thin = Side(border_style="thin", color="888888")
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            if cell.row == 1:
                cell.font = Font(bold=True); cell.fill = header_fill
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_length + 2, 15)

    # 2b) Định dạng cột "Log in date" và "ETD" về dd-mmm nếu chưa đúng
    format_date_columns(ws, target_headers=("log in date", "etd"), number_format="DD-MMM")

    # 3) Ẩn dòng theo Login Date (nếu có cột đó)
    if login_col_trf:
        hide_rows_by_login_date(ws, login_col_trf, today=datetime.now().date())

    wb.save(trf_file)
    print("Da xuat file:", trf_file)

    # 4) Upload lên SharePoint
    upload_relative_url_trf = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/TRF.xlsx"
    folder_excel_trf = os.path.dirname(upload_relative_url_trf)
    ensure_folder(ctx, folder_excel_trf)
    with open(trf_file, "rb") as f:
        ctx.web.get_folder_by_server_relative_url(folder_excel_trf)\
           .upload_file(os.path.basename(upload_relative_url_trf), f.read()).execute_query()
    print("Da upload file", trf_file, "len SharePoint:", upload_relative_url_trf)
