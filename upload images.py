import os
import zipfile
import hashlib
import json
import time
from datetime import datetime
from PIL import Image
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# ==== CONFIG ====
site_url = "https://jonathancharles.sharepoint.com/sites/TESTLAB-VFR9"
username = "tan_qa@vfr.net.vn"
password = "qaz@Tat@123"
upload_folder_sharepoint = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/IMAGES_ZIP/"
local_images = "images"
MAX_ZIP_SIZE_MB = 200
MAX_ZIP_SIZE = MAX_ZIP_SIZE_MB * 1024 * 1024
tmp_download_folder = "__tmp_sharepoint__"
if not os.path.exists(tmp_download_folder):
    os.makedirs(tmp_download_folder)

# ==== Tiện ích dọn file .txt rác: repeat / notified, chỉ xóa khi ≥ 2 ngày ====
def clean_noise_txt(root_folder: str, dry_run: bool = False, min_age_days: int = 2) -> int:
    """
    Xóa các file .txt có 'repeat' hoặc 'notified' trong tên (không phân biệt hoa/thường)
    CHỈ KHI file đã tồn tại ít nhất `min_age_days` ngày (dựa trên mtime).

    - root_folder: thư mục gốc để quét
    - dry_run: True -> chỉ in ra file sẽ xóa, không xóa thật
    - min_age_days: số ngày tối thiểu file phải "già" để được xóa

    Trả về: số lượng file đã (hoặc sẽ) xóa.
    """
    keywords = ("repeat", "notified")
    deleted = 0
    if not os.path.isdir(root_folder):
        return 0

    now = time.time()
    threshold_seconds = min_age_days * 24 * 60 * 60

    for foldername, _, filenames in os.walk(root_folder):
        for filename in filenames:
            name_lower = filename.lower()
            if not (name_lower.endswith(".txt") and any(kw in name_lower for kw in keywords)):
                continue

            file_path = os.path.join(foldername, filename)
            try:
                mtime = os.path.getmtime(file_path)  # thời điểm sửa đổi cuối cùng
            except FileNotFoundError:
                # Có thể file vừa bị xóa/di chuyển
                continue
            age_seconds = now - mtime

            if age_seconds >= threshold_seconds:
                if dry_run:
                    print(f"[DRY-RUN] Sẽ xóa: {file_path} (tuổi ~ {age_seconds/86400:.2f} ngày)")
                    deleted += 1
                else:
                    try:
                        os.remove(file_path)
                        print(f"Đã xóa: {file_path} (tuổi ~ {age_seconds/86400:.2f} ngày)")
                        deleted += 1
                    except Exception as e:
                        print(f"Lỗi khi xóa {file_path}: {e}")
            else:
                # File còn mới (< 2 ngày), không xóa
                print(f"Giữ lại (mới < {min_age_days} ngày): {file_path}")
    return deleted

# ==== Hàm nén ảnh trực tiếp (chỉ khi cần) ====
def compress_image_inplace(path, quality=70, max_size=(1920,1080)):
    try:
        img = Image.open(path)
        img.thumbnail(max_size, Image.LANCZOS)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.save(path, "JPEG", quality=quality, optimize=True)
    except Exception as e:
        print(f"Lỗi nén ảnh {path}: {e}")

# ==== Metadata: Lưu/đọc thông tin nén trong từng folder con ====
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

def compress_folder_inplace_smart(folder, quality=70, max_size=(1920,1080)):
    old_meta = read_compress_meta(folder)
    files_info = {}
    nened = False
    for root, _, files in os.walk(folder):
        for file in files:
            ext = file.lower().split(".")[-1]
            if ext not in ["jpg", "jpeg", "png"]:
                continue
            path = os.path.join(root, file)
            mtime = os.path.getmtime(path)
            size = os.path.getsize(path)
            key = os.path.relpath(path, folder)
            files_info[key] = [mtime, size]
            old = old_meta.get(key)
            # Nếu file mới, hoặc đã sửa/đổi size, mới nén lại!
            if (old is None) or (old[0] != mtime or old[1] != size):
                compress_image_inplace(path, quality, max_size)
                nened = True
                # Cập nhật lại mtime và size sau nén
                mtime2 = os.path.getmtime(path)
                size2 = os.path.getsize(path)
                files_info[key] = [mtime2, size2]
    write_compress_meta(folder, files_info)
    return nened

def md5sum(filename, bufsize=65536):
    h = hashlib.md5()
    with open(filename, 'rb') as f:
        while True:
            chunk = f.read(bufsize)
            if not chunk:
                break
            h.update(chunk)
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

def folder_month(folder):
    """Lấy tháng theo file earliest created trong folder."""
    min_time = None
    for root, _, files in os.walk(folder):
        for file in files:
            path = os.path.join(root, file)
            ctime = os.path.getctime(path)
            if min_time is None or ctime < min_time:
                min_time = ctime
    if not min_time:
        return None
    return datetime.fromtimestamp(min_time).strftime('%Y%m')

def get_folder_size(folder):
    total = 0
    for dirpath, _, filenames in os.walk(folder):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.isfile(fp):
                total += os.path.getsize(fp)
    return total

# ==== Đăng nhập SharePoint ====
ctx_auth = AuthenticationContext(site_url)
if not ctx_auth.acquire_token_for_user(username, password):
    raise Exception("Không kết nối được SharePoint!")
ctx = ClientContext(site_url, ctx_auth)

# ==== Duyệt từng folder, dọn .txt rác (≥2 ngày), nén nếu cần, ghi log ====
folders = [os.path.join(local_images, f) for f in os.listdir(local_images) if os.path.isdir(os.path.join(local_images, f))]
folders.sort()

folders_by_month = {}
for folder in folders:
    folder_name = os.path.basename(folder)

    # 1) Dọn file .txt rác trước (chỉ xóa khi ≥ 2 ngày)
    print(f"Clean .txt (repeat/notified, tuổi ≥ 2 ngày) trong folder {folder_name} ...")
    removed_count = clean_noise_txt(folder, dry_run=False, min_age_days=2)
    if removed_count:
        print(f"→ Đã xóa {removed_count} file .txt rác.")
    else:
        print("→ Không có file .txt rác đủ điều kiện để xóa.")

    # 2) Nén ảnh thông minh
    print(f"Check nén folder {folder_name} ...")
    nened = compress_folder_inplace_smart(folder, quality=70, max_size=(1920,1080))
    if nened:
        print(f"→ Đã nén lại ảnh mới/chưa nén.")
    else:
        print(f"→ Không có ảnh mới, bỏ qua nén.")

    # 3) Gom nhóm theo tháng
    thang = folder_month(folder)
    if not thang:
        continue
    if thang not in folders_by_month:
        folders_by_month[thang] = []
    folders_by_month[thang].append(folder)

# ==== Gom các folder cùng tháng thành 1 zip (<= 200MB/zip) ====
for thang, month_folders in sorted(folders_by_month.items()):
    cur_group = []
    cur_size = 0
    group_idx = 1
    for folder in month_folders:
        fsize = get_folder_size(folder)
        if cur_group and cur_size + fsize > MAX_ZIP_SIZE:
            zip_file = f"images_{thang}_part{group_idx}.zip" if group_idx > 1 else f"images_{thang}.zip"
            # Tạo zip
            with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for f in cur_group:
                    for root, _, files in os.walk(f):
                        for file in files:
                            full_path = os.path.join(root, file)
                            arcname = os.path.relpath(full_path, os.path.dirname(month_folders[0]))
                            zipf.write(full_path, arcname)
            # Check hash với SharePoint
            local_md5 = md5sum(zip_file)
            remote_zip_path = os.path.join(tmp_download_folder, zip_file)
            exists = download_file_from_sharepoint(ctx, upload_folder_sharepoint, zip_file, remote_zip_path)
            if exists:
                remote_md5 = md5sum(remote_zip_path)
                if remote_md5 == local_md5:
                    print(f"Không có thay đổi trong {zip_file}, KHÔNG upload lên SharePoint.")
                    os.remove(zip_file)
                    os.remove(remote_zip_path)
                    cur_group = []
                    cur_size = 0
                    group_idx += 1
                    continue
                else:
                    print(f"{zip_file} trên SharePoint đã khác, sẽ upload mới (ghi đè).")
                os.remove(remote_zip_path)
            else:
                print(f"SharePoint chưa có {zip_file}, sẽ upload mới.")

            with open(zip_file, "rb") as fz:
                ctx.web.get_folder_by_server_relative_url(upload_folder_sharepoint) \
                    .upload_file(zip_file, fz.read()).execute_query()
            print(f"Đã upload {zip_file} lên SharePoint.")
            os.remove(zip_file)
            print(f"Đã xoá {zip_file} ở local.")

            # Reset group
            cur_group = []
            cur_size = 0
            group_idx += 1
        cur_group.append(folder)
        cur_size += fsize
    # Xử lý phần cuối cùng chưa zip
    if cur_group:
        zip_file = f"images_{thang}_part{group_idx}.zip" if group_idx > 1 else f"images_{thang}.zip"
        with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for f in cur_group:
                for root, _, files in os.walk(f):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, os.path.dirname(month_folders[0]))
                        zipf.write(full_path, arcname)
        local_md5 = md5sum(zip_file)
        remote_zip_path = os.path.join(tmp_download_folder, zip_file)
        exists = download_file_from_sharepoint(ctx, upload_folder_sharepoint, zip_file, remote_zip_path)
        if exists:
            remote_md5 = md5sum(remote_zip_path)
            if remote_md5 == local_md5:
                print(f"Không có thay đổi trong {zip_file}, KHÔNG upload lên SharePoint.")
                os.remove(zip_file)
                os.remove(remote_zip_path)
                continue
            else:
                print(f"{zip_file} trên SharePoint đã khác, sẽ upload mới (ghi đè).")
            os.remove(remote_zip_path)
        else:
            print(f"SharePoint chưa có {zip_file}, sẽ upload mới.")

        with open(zip_file, "rb") as fz:
            ctx.web.get_folder_by_server_relative_url(upload_folder_sharepoint) \
                .upload_file(zip_file, fz.read()).execute_query()
        print(f"Đã upload {zip_file} lên SharePoint.")
        os.remove(zip_file)
        print(f"Đã xoá {zip_file} ở local.")

print("\nĐã xử lý xong tất cả các nhóm folder theo tháng (đã dọn .txt rác cũ ≥2 ngày, ảnh chỉ nén và up khi có thay đổi).")
