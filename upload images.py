import os
import zipfile
import hashlib
import json
import time
import shutil
import tempfile
import re
from datetime import datetime, date
from collections import Counter
from PIL import Image, ImageOps, ExifTags

# Office365 SharePoint CSOM
from office365.sharepoint.client_context import ClientContext

# MSAL for OAuth (Delegated + cache), same pattern as "excel export.py"
import msal

# ================== OAUTH CONFIG (DELEGATED + CACHE) ==================
TENANT_ID = "064944f6-1e04-4050-b3e1-e361758625ec"       # Directory (tenant) ID
CLIENT_ID = "9abf6ee2-50c8-47c8-a9f2-8cf18587c6ea"       # Application (client) ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SP_HOST = "https://jonathancharles.sharepoint.com"

# SharePoint delegated scopes (match "excel export.py")
SPO_SCOPES = [
    f"{SP_HOST}/AllSites.Read",
    f"{SP_HOST}/AllSites.Write",
]

TOKEN_CACHE_FILE = "token_cache.bin"

# ==== SITE & PATH CONFIG ====
site_url = "https://jonathancharles.sharepoint.com/sites/TESTLAB-VFR9"
upload_folder_sharepoint = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/IMAGES_ZIP/"
local_images = "images"

# (legacy leftover; kept harmless)
tmp_download_folder = "__tmp_sharepoint__"
os.makedirs(tmp_download_folder, exist_ok=True)

# =============== TXT CLEAN ===============
def clean_unlabeled_txt(root_folder: str, dry_run: bool = False, min_age_days: int = 0) -> int:
    if not os.path.isdir(root_folder):
        return 0
    keep_keywords = ("comment", "status")
    deleted = 0
    now = time.time()
    threshold_seconds = max(0, min_age_days) * 24 * 60 * 60
    for foldername, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if not filename.lower().endswith(".txt"):
                continue
            name_lower = filename.lower()
            if any(kw in name_lower for kw in keep_keywords):
                continue
            file_path = os.path.join(foldername, filename)
            try:
                mtime = os.path.getmtime(file_path)
            except FileNotFoundError:
                continue
            age_seconds = now - mtime
            if threshold_seconds > 0 and age_seconds < threshold_seconds:
                continue
            if dry_run:
                print(f"[DRY-RUN] Sẽ xoá: {file_path}")
                deleted += 1
            else:
                try:
                    os.remove(file_path)
                    print(f"Đã xoá: {file_path}")
                    deleted += 1
                except Exception as e:
                    print(f"Lỗi khi xoá {file_path}: {e}")
    return deleted

# =============== IMAGE COMPRESS (IN-PLACE, IDEMPOTENT) ===============
def compress_image_inplace(path, quality=80, max_side=2000):
    try:
        with Image.open(path) as img:
            img = ImageOps.exif_transpose(img)
            w, h = img.size
            m = max(w, h)
            if m > max_side:
                scale = max_side / float(m)
                img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            img.save(path, "JPEG", quality=quality, optimize=True)
    except Exception as e:
        print(f"[WARN] Lỗi nén ảnh {path}: {e}")

def write_compress_meta(folder, files_info):
    meta_file = os.path.join(folder, ".compressed_info.json")
    with open(meta_file, "w") as f:
        json.dump(files_info, f)

def read_compress_meta(folder):
    meta_file = os.path.join(folder, ".compressed_info.json")
    if not os.path.exists(meta_file):
        return {}
    with open(meta_file, "r") as f:
        return json.load(f)

def compress_folder_inplace_smart(folder, quality=80, max_side=2000):
    old_meta = read_compress_meta(folder)
    files_info = {}
    changed = False
    for root, _, files in os.walk(folder):
        for file in files:
            ext = file.lower().split(".")[-1]
            if ext not in ["jpg", "jpeg", "png"]:
                continue
            path = os.path.join(root, file)
            try:
                mtime = os.path.getmtime(path)
                size = os.path.getsize(path)
            except FileNotFoundError:
                continue
            key = os.path.relpath(path, folder)
            files_info[key] = [mtime, size]
            old = old_meta.get(key)
            if (old is None) or (old[0] != mtime or old[1] != size):
                compress_image_inplace(path, quality=80, max_side=2000)
                changed = True
                try:
                    mtime2 = os.path.getmtime(path)
                    size2 = os.path.getsize(path)
                    files_info[key] = [mtime2, size2]
                except FileNotFoundError:
                    pass
    write_compress_meta(folder, files_info)
    return changed

# =============== SHAREPOINT & ZIP UTILS ===============
def md5sum(filename, bufsize=65536):
    h = hashlib.md5()
    with open(filename, "rb") as f:
        while True:
            chunk = f.read(bufsize)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()

def file_md5(path, bufsize=65536):
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            b = f.read(bufsize)
            if not b:
                break
            h.update(b)
    return h.hexdigest()

def download_file_from_sharepoint(ctx, folder_url, filename, local_path):
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    files = folder.files
    ctx.load(files)
    ctx.execute_query()
    for f in files:
        if f.properties["Name"] == filename:
            with open(local_path, "wb") as out:
                f.download(out).execute_query()
            return True
    return False

def upload_file_to_sharepoint(ctx, folder_url, local_file, remote_name=None, chunk_size=10 * 1024 * 1024):
    """
    Upload file lên SharePoint:
      - < 50MB: upload thẳng (Files.add)
      - >= 50MB: dùng create_upload_session (chunked)
    """
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)

    src_path = local_file
    desired_name = remote_name or os.path.basename(local_file)
    file_size = os.path.getsize(local_file)

    # Nhỏ hơn 50MB: giữ cách cũ
    if file_size < 50 * 1024 * 1024:
        with open(local_file, "rb") as fz:
            folder.upload_file(desired_name, fz.read()).execute_query()
        return

    # Lớn: dùng upload session (chunked)
    # Nếu muốn đổi tên file trên SharePoint, tạo bản tạm để giữ đúng basename
    temp_to_upload = None
    try:
        if desired_name != os.path.basename(local_file):
            temp_dir = tempfile.mkdtemp(prefix="sp_upload_")
            temp_to_upload = os.path.join(temp_dir, desired_name)
            shutil.copy2(local_file, temp_to_upload)
            src_path = temp_to_upload

        # create_upload_session sẽ tự chia chunk và upload
        # chunk_size mặc định 10MB; có thể tăng 20–60MB nếu băng thông ổn
        uploaded_file = folder.files.create_upload_session(src_path, chunk_size=chunk_size).execute_query()
        # uploaded_file sau khi execute_query() là đối tượng File trên SharePoint
        # Không cần làm gì thêm
    finally:
        if temp_to_upload and os.path.exists(os.path.dirname(temp_to_upload)):
            shutil.rmtree(os.path.dirname(temp_to_upload), ignore_errors=True)

def unzip_to_dir(zip_path, dest_dir):
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(dest_dir)

# =============== DATE DERIVATION FROM IMAGES ===============
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp", ".heic", ".heif"}

def is_image_file(path):
    return os.path.splitext(path)[1].lower() in IMAGE_EXTS

def parse_date_from_name(name: str):
    s = name.lower()
    pats = [
        r'(?P<y>20\d{2})[-_\.]?(?P<m>0[1-9]|1[0-2])[-_\.]?(?P<d>0[1-9]|[12]\d|3[01])',  # 20250817 / 2025-08-17
        r'(?P<d>0[1-9]|[12]\d|3[01])[-_\.](?P<m>0[1-9]|1[0-2])[-_\.](?P<y>20\d{2})',     # 17-08-2025
    ]
    for pat in pats:
        m = re.search(pat, s)
        if m:
            try:
                y, mth, d = int(m.group("y")), int(m.group("m")), int(m.group("d"))
                return datetime(y, mth, d)
            except Exception:
                continue
    return None

EXIF_KEYS = {k for k, v in ExifTags.TAGS.items() if v in ("DateTimeOriginal", "DateTimeDigitized", "DateTime")}

def get_exif_datetime(path):
    try:
        with Image.open(path) as im:
            exif = getattr(im, "getexif", lambda: None)()
            if not exif:
                return None
            for k in EXIF_KEYS:
                if k in exif:
                    val = str(exif.get(k))
                    try:
                        return datetime.strptime(val.replace("-", ":").replace(".", ":"), "%Y:%m:%d %H:%M:%S")
                    except Exception:
                        pass
    except Exception:
        pass
    return None

def best_guess_datetime(path):
    dt = get_exif_datetime(path)
    if dt:
        return dt
    dt = parse_date_from_name(os.path.basename(path))
    if dt:
        return dt
    try:
        return datetime.fromtimestamp(os.path.getmtime(path))
    except Exception:
        return None

def folder_day_by_images(folder) -> str:
    """
    Tính ngày cho CẢ folder dựa trên ảnh bên trong:
      - Lấy ngày (YYYY-MM-DD) cho TỪNG ảnh: EXIF -> tên file -> mtime.
      - Chọn NGÀY PHỔ BIẾN NHẤT (mode). Nếu hoà, lấy ngày SỚM NHẤT.
      - Nếu không có ảnh → dùng earliest ctime của mọi file.
    """
    dates: list[date] = []
    earliest_ts = None

    for root, _, files in os.walk(folder):
        for fname in files:
            full = os.path.join(root, fname)
            # track earliest ts as fallback
            try:
                ctime = os.path.getctime(full)
                earliest_ts = ctime if earliest_ts is None or ctime < earliest_ts else earliest_ts
            except FileNotFoundError:
                pass

            if not is_image_file(full):
                continue
            dt = best_guess_datetime(full)
            if dt:
                dates.append(dt.date())

    if dates:
        cnt = Counter(dates)
        max_count = max(cnt.values())
        candidates = [d for d, c in cnt.items() if c == max_count]
        chosen = min(candidates)  # tie-breaker: earliest
        return chosen.strftime("%Y-%m-%d")

    if earliest_ts is not None:
        return datetime.fromtimestamp(earliest_ts).strftime("%Y-%m-%d")

    # last resort: today
    return datetime.now().strftime("%Y-%m-%d")

def month_from_day_label(day_label: str) -> str:
    # "YYYY-MM-DD" -> "YYYYMM"
    try:
        return day_label[:4] + day_label[5:7]
    except Exception:
        return datetime.now().strftime("%Y%m")

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def copy_merge_folder(src_folder, dst_folder):
    """
    Gộp src_folder vào dst_folder:
      - bỏ .txt không nhãn
      - nếu file chưa tồn tại: copy
      - nếu tồn tại: so sánh size+md5; nếu khác -> ghi đè
    """
    for root, _, files in os.walk(src_folder):
        for fname in files:
            src_path = os.path.join(root, fname)
            if fname.lower().endswith(".txt"):
                low = fname.lower()
                if ("comment" not in low) and ("status" not in low):
                    continue
            rel = os.path.relpath(src_path, src_folder)
            dst_path = os.path.join(dst_folder, rel)
            ensure_dir(os.path.dirname(dst_path))
            if not os.path.exists(dst_path):
                shutil.copy2(src_path, dst_path)
            else:
                try:
                    if os.path.getsize(src_path) != os.path.getsize(dst_path) or file_md5(src_path) != file_md5(dst_path):
                        shutil.copy2(src_path, dst_path)
                except FileNotFoundError:
                    shutil.copy2(src_path, dst_path)

# ================== MSAL TOKEN CACHE HELPERS ==================
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
    """
    ctx = ClientContext(site_url)

    def _auth(request):
        token = acquire_spo_access_token()  # lấy từ cache, tự refresh khi cần
        request.ensure_header("Authorization", "Bearer " + token)

    # Ghi đè cơ chế auth mặc định
    ctx.authentication_context.authenticate_request = _auth
    print("AUTH MODE:", "Delegated (Device Code + Token Cache) [custom auth]")
    return ctx

# =============== MAIN PIPELINE ===============
def process():
    ctx = get_ctx(site_url)

    # Liệt kê các folder 25-xxxx
    if not os.path.isdir(local_images):
        print(f"[WARN] Không tìm thấy thư mục local_images: {local_images}")
        return

    folders = [os.path.join(local_images, f) for f in os.listdir(local_images)
               if os.path.isdir(os.path.join(local_images, f))]
    folders.sort()

    # 1) Clean + Resize in-place trên GỐC (idempotent)
    for folder in folders:
        clean_unlabeled_txt(folder, dry_run=False, min_age_days=0)
        compress_folder_inplace_smart(folder, quality=80, max_side=2000)

    # 2) Gom bucket theo THÁNG, nhưng tháng lấy từ day_label (của folder theo ảnh)
    month_buckets = {}  # { 'YYYYMM': [(folder_path, day_label)] }
    for folder in folders:
        day_label = folder_day_by_images(folder)      # YYYY-MM-DD (từ ảnh)
        m_label = month_from_day_label(day_label)     # YYYYMM
        month_buckets.setdefault(m_label, []).append((folder, day_label))

    # 3) Với mỗi tháng -> tải zip cũ (nếu có), merge, thêm local, zip lại & upload
    for thang, entries in sorted(month_buckets.items()):
        month_zip_name = f"images_{thang}.zip"

        with tempfile.TemporaryDirectory() as staging_parent:
            staging_root = os.path.join(staging_parent, f"{thang}_root")  # YYYY-MM-DD/25-xxxx/...
            ensure_dir(staging_root)

            # 3.1 Merge từ SharePoint nếu có
            sp_local_copy = os.path.join(staging_parent, f"dl_{month_zip_name}")
            exists_remote = download_file_from_sharepoint(ctx, upload_folder_sharepoint, month_zip_name, sp_local_copy)
            if exists_remote:
                print(f"[INFO] Tải {month_zip_name} từ SharePoint để merge...")
                unzip_to_dir(sp_local_copy, staging_parent)

                # Tìm tất cả thư mục ngày đã unzip, copy vào staging_root
                # (không giả định exact root trong zip; copy mọi dir con vào staging_root)
                for item in os.listdir(staging_parent):
                    p = os.path.join(staging_parent, item)
                    if os.path.isdir(p) and item != os.path.basename(staging_root):
                        for day_folder in os.listdir(p):
                            src_day = os.path.join(p, day_folder)
                            dst_day = os.path.join(staging_root, day_folder)
                            if os.path.isdir(src_day):
                                ensure_dir(dst_day)
                                copy_merge_folder(src_day, dst_day)

            # 3.2 Thêm dữ liệu LOCAL theo đúng ngày
            for folder, day_label in entries:
                dst_day_dir = os.path.join(staging_root, day_label)
                dst_folder = os.path.join(dst_day_dir, os.path.basename(folder))
                ensure_dir(dst_folder)
                copy_merge_folder(folder, dst_folder)

            # 3.3 Đóng gói lại zip tháng (không giữ byday trên máy)
            parent_for_zip = os.path.dirname(staging_root)
            zip_path_temp = os.path.join(staging_parent, month_zip_name)
            with zipfile.ZipFile(zip_path_temp, "w", zipfile.ZIP_DEFLATED) as zf:
                for root, _, files in os.walk(staging_root):
                    for name in files:
                        full = os.path.join(root, name)
                        arc = os.path.relpath(full, parent_for_zip)
                        zf.write(full, arc)

            # 3.4 So sánh MD5 với bản cũ & upload nếu có thay đổi
            if exists_remote:
                old_md5 = md5sum(sp_local_copy)
                new_md5 = md5sum(zip_path_temp)
                if old_md5 == new_md5:
                    print(f"[INFO] {month_zip_name}: Không có thay đổi. Bỏ qua upload.")
                else:
                    print(f"[INFO] {month_zip_name}: Có thay đổi. Upload lên SharePoint...")
                    upload_file_to_sharepoint(ctx, upload_folder_sharepoint, zip_path_temp, month_zip_name)
                    print(f"[OK] Uploaded {month_zip_name}")
            else:
                print(f"[INFO] {month_zip_name} chưa tồn tại. Upload mới...")
                upload_file_to_sharepoint(ctx, upload_folder_sharepoint, zip_path_temp, month_zip_name)
                print(f"[OK] Uploaded {month_zip_name}")

if __name__ == "__main__":
    process()
    print("\nHoàn tất: gom theo NGÀY dựa trên metadata ảnh (không dùng tên folder), merge với zip tháng trên SharePoint, zip lại & upload. Không giữ byday/zip ở local.")
