import os
import zipfile
import hashlib
import json
import time
import shutil
import tempfile
from datetime import datetime
from PIL import Image, ImageOps
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# ==== CONFIG ====
site_url = "https://jonathancharles.sharepoint.com/sites/TESTLAB-VFR9"
username = "tan_qa@vfr.net.vn"
password = "qaz@Tat@123"
upload_folder_sharepoint = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/IMAGES_ZIP/"
local_images = "images"

# Nếu sau này muốn quay lại split theo kích thước, bổ sung logic partN.
# Hiện tại: mỗi tháng **một** zip: images_YYYYMM.zip
# MAX_ZIP_SIZE_MB = 200

tmp_download_folder = "__tmp_sharepoint__"
os.makedirs(tmp_download_folder, exist_ok=True)

# ==== 1) Tiện ích dọn .txt: xoá .txt KHÔNG có 'comment' hoặc 'status' trong tên ====
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
                # giữ .txt có 'comment' hoặc 'status'
                continue
            file_path = os.path.join(foldername, filename)
            try:
                mtime = os.path.getmtime(file_path)
            except FileNotFoundError:
                continue
            age_seconds = now - mtime
            if threshold_seconds > 0 and age_seconds < threshold_seconds:
                # còn mới, chưa xoá nếu có yêu cầu tuổi file
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

# ==== 2) Resize ảnh gốc **in-place** (idempotent) ====
def compress_image_inplace(path, quality=80, max_side=2000):
    try:
        with Image.open(path) as img:
            img = ImageOps.exif_transpose(img)
            # downscale nếu cần
            w, h = img.size
            m = max(w, h)
            if m > max_side:
                scale = max_side / float(m)
                img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
            # chuẩn hóa màu cho JPEG
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            # lưu lại (JPEG) tối ưu
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
                # chỉ nén khi mới hoặc thay đổi
                compress_image_inplace(path, quality=quality, max_side=max_side)
                changed = True
                # cập nhật metadata sau nén
                try:
                    mtime2 = os.path.getmtime(path)
                    size2 = os.path.getsize(path)
                    files_info[key] = [mtime2, size2]
                except FileNotFoundError:
                    pass
    write_compress_meta(folder, files_info)
    return changed

# ==== 3) Tiện ích SharePoint & ZIP ====
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

def upload_file_to_sharepoint(ctx, folder_url, local_file, remote_name=None):
    if remote_name is None:
        remote_name = os.path.basename(local_file)
    with open(local_file, "rb") as fz:
        ctx.web.get_folder_by_server_relative_url(folder_url) \
            .upload_file(remote_name, fz.read()).execute_query()

def unzip_to_dir(zip_path, dest_dir):
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(dest_dir)

def zip_dir(src_dir, zip_path):
    # nén toàn bộ src_dir vào zip_path; arcname bắt đầu từ src_dir basename
    base_parent = os.path.dirname(src_dir)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(src_dir):
            for name in files:
                full = os.path.join(root, name)
                arc = os.path.relpath(full, base_parent)
                zf.write(full, arc)

# ==== 4) Phân nhóm theo THÁNG & NGÀY ở cấp folder 25-xxxx ====
def folder_month(folder):
    """
    Lấy YYYYMM: ưu tiên theo earliest ctime của file trong folder.
    """
    min_time = None
    for root, _, files in os.walk(folder):
        for file in files:
            p = os.path.join(root, file)
            try:
                ctime = os.path.getctime(p)
            except FileNotFoundError:
                continue
            if min_time is None or ctime < min_time:
                min_time = ctime
    if not min_time:
        # fallback: nếu không có file, lấy theo ngày từ tên folder + tháng hiện tại
        # (trường hợp hiếm)
        return datetime.now().strftime("%Y%m")
    return datetime.fromtimestamp(min_time).strftime("%Y%m")

def folder_day_from_name(folder, month_label):
    """
    Từ tên '25-xxxx' suy ra YYYY-MM-DD. Nếu không parse được -> 'YYYY-MM-unknown'.
    """
    name = os.path.basename(folder)
    prefix = name.split("-")[0]
    try:
        day = int(prefix)
        year = month_label[:4]
        month = month_label[4:6]
        return f"{year}-{month}-{day:02d}"
    except:
        return f"{month_label[:4]}-{month_label[4:6]}-unknown"

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def copy_merge_folder(src_folder, dst_folder):
    """
    Gộp src_folder vào dst_folder:
      - nếu file chưa tồn tại: copy
      - nếu tồn tại: so sánh md5; nếu khác -> ghi đè
    """
    for root, _, files in os.walk(src_folder):
        for fname in files:
            src_path = os.path.join(root, fname)
            # bỏ .txt không nhãn (an toàn thêm lần nữa trong bước staging)
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

# ==== 5) Đăng nhập SharePoint ====
ctx_auth = AuthenticationContext(site_url)
if not ctx_auth.acquire_token_for_user(username, password):
    raise Exception("Không kết nối được SharePoint!")
ctx = ClientContext(site_url, ctx_auth)

# ==== 6) Pipeline chính ====
def process():
    # 6.1 Liệt kê các folder 25-xxxx trong images/
    month_buckets = {}  # { 'YYYYMM': [folder_abs_paths...] }
    folders = [os.path.join(local_images, f) for f in os.listdir(local_images)
               if os.path.isdir(os.path.join(local_images, f))]
    folders.sort()

    for folder in folders:
        folder_name = os.path.basename(folder)

        # (A) CLEAN TXT in-place (idempotent)
        clean_unlabeled_txt(folder, dry_run=False, min_age_days=0)

        # (B) RESIZE in-place (idempotent)
        compress_folder_inplace_smart(folder, quality=80, max_side=2000)

        # (C) BUCKET theo tháng
        thang = folder_month(folder)
        month_buckets.setdefault(thang, []).append(folder)

    # 6.2 Với mỗi THÁNG -> build zip theo ngày/folder
    for thang, month_folders in sorted(month_buckets.items()):
        month_zip_name = f"images_{thang}.zip"
        month_zip_local = os.path.join(tempfile.gettempdir(), month_zip_name)  # file zip tạm

        # Tạo một thư mục staging tạm để build/merge (sẽ xoá cuối)
        with tempfile.TemporaryDirectory() as staging_parent:
            staging_root = os.path.join(staging_parent, f"{thang}_root")  # root chứa YYYY-MM-DD/25-xxxx
            ensure_dir(staging_root)

            # 6.2.1 Nếu SharePoint đã có zip tháng -> tải về & giải nén vào staging_root để MERGE
            sp_local_copy = os.path.join(staging_parent, f"dl_{month_zip_name}")
            exists_remote = download_file_from_sharepoint(ctx, upload_folder_sharepoint, month_zip_name, sp_local_copy)
            if exists_remote:
                print(f"[INFO] Đã tìm thấy {month_zip_name} trên SharePoint. Tải về để merge...")
                unzip_to_dir(sp_local_copy, staging_parent)
                # sau unzip, root sẽ là cùng cấp với tên thư mục gốc trong zip (do zip_dir đặt relative)
                # ta di chuyển mọi thứ dưới staging_parent vào staging_root nếu cần
                # Tìm thư mục chứa các ngày:
                candidates = []
                for item in os.listdir(staging_parent):
                    p = os.path.join(staging_parent, item)
                    if os.path.isdir(p) and item != os.path.basename(staging_root):
                        candidates.append(p)
                # Copy tất cả folder con vào staging_root
                for c in candidates:
                    for day_folder in os.listdir(c):
                        src_day = os.path.join(c, day_folder)
                        dst_day = os.path.join(staging_root, day_folder)
                        if os.path.isdir(src_day):
                            ensure_dir(dst_day)
                            copy_merge_folder(src_day, dst_day)

            # 6.2.2 Từ local, gộp các folder 25-xxxx vào **đúng ngày** trong staging_root
            for folder in month_folders:
                day_label = folder_day_from_name(folder, thang)   # YYYY-MM-DD
                dst_day_dir = os.path.join(staging_root, day_label)
                dst_folder = os.path.join(dst_day_dir, os.path.basename(folder))
                ensure_dir(dst_folder)
                copy_merge_folder(folder, dst_folder)

            # 6.2.3 Nén staging_root thành tháng.zip
            #     (Không giữ byday trên máy: staging_root là tạm, sẽ xoá)
            # Để có cấu trúc: images_YYYYMM.zip/ YYYY-MM-DD/25-xxxx/...
            # ta nén cha của staging_root để giữ tên thư mục cấp 1
            parent_for_zip = os.path.dirname(staging_root)
            zip_path_temp = os.path.join(staging_parent, month_zip_name)
            with zipfile.ZipFile(zip_path_temp, "w", zipfile.ZIP_DEFLATED) as zf:
                for root, _, files in os.walk(staging_root):
                    for name in files:
                        full = os.path.join(root, name)
                        arc = os.path.relpath(full, parent_for_zip)
                        zf.write(full, arc)

            # 6.2.4 So sánh với bản cũ trên SharePoint (nếu có). Nếu khác -> upload; nếu giống -> bỏ qua
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

            # 6.2.5 Không giữ zip cục bộ và không giữ byday: staging_parent sẽ tự xoá (tempdir)
            # (Không cần os.remove(zip_path_temp) vì thuộc staging_parent; nhưng nếu đổi đường dẫn thì gọi os.remove)

if __name__ == "__main__":
    process()
    print("\nHoàn tất: mỗi lần chạy đều đọc zip tháng trên SharePoint (nếu có), merge dữ liệu LOCAL mới theo NGÀY/25-xxxx, zip lại & upload. Không giữ byday/zip ở local.")
