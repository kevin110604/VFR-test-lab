import os
import zipfile
import hashlib
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# ==== Cấu hình ====
site_url = "https://jonathancharles.sharepoint.com/sites/TESTLAB-VFR9"
username = "tan_qa@vfr.net.vn"
password = "qaz@Tat@123"
upload_folder_sharepoint = "/sites/TESTLAB-VFR9/Shared Documents/DATA DAILY/IMAGES_ZIP/"
local_images = "images"
MAX_ZIP_SIZE_MB = 300  # Mỗi file zip tối đa ~200MB
MAX_ZIP_SIZE = MAX_ZIP_SIZE_MB * 1024 * 1024

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
            if "already exists" not in str(e).lower():
                print(f"Lỗi tạo folder {current_url}: {e}")
                raise
    return ctx.web.get_folder_by_server_relative_url(folder_url)

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

def zip_folders_by_size(folder_list, zip_name):
    """Nén nhiều folder con thành 1 file zip (giữ nguyên cấu trúc)."""
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for folder in folder_list:
            for root, _, files in os.walk(folder):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, os.path.dirname(folder_list[0]))
                    zipf.write(full_path, arcname)

def get_folder_size(folder):
    total = 0
    for dirpath, dirnames, filenames in os.walk(folder):
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

ensure_folder(ctx, upload_folder_sharepoint)
tmp_download_folder = "__tmp_sharepoint__"
if not os.path.exists(tmp_download_folder):
    os.makedirs(tmp_download_folder)

# ==== Gom nhóm folder con thành từng group theo dung lượng tối đa ====
folders = [os.path.join(local_images, folder) for folder in os.listdir(local_images) if os.path.isdir(os.path.join(local_images, folder))]
folders.sort()  # Đảm bảo thứ tự

groups = []
cur_group = []
cur_size = 0
for folder in folders:
    fsize = get_folder_size(folder)
    if cur_group and cur_size + fsize > MAX_ZIP_SIZE:
        groups.append(cur_group)
        cur_group = []
        cur_size = 0
    cur_group.append(folder)
    cur_size += fsize
if cur_group:
    groups.append(cur_group)

print(f"Tổng số nhóm zip sẽ tạo: {len(groups)} (mỗi nhóm <= {MAX_ZIP_SIZE_MB} MB)")

for idx, group in enumerate(groups, 1):
    folders_in_group = [os.path.basename(f) for f in group]
    zip_file = f"images_group{idx}.zip"
    # Xoá zip cũ nếu có
    if os.path.exists(zip_file):
        os.remove(zip_file)
    print(f"\nĐang nén group {idx} gồm các folder: {folders_in_group} thành {zip_file} ...")
    zip_folders_by_size(group, zip_file)
    local_md5 = md5sum(zip_file)

    # Kiểm tra zip này đã có trên SharePoint chưa
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

    # Upload zip lên SharePoint
    with open(zip_file, "rb") as f:
        ctx.web.get_folder_by_server_relative_url(upload_folder_sharepoint) \
            .upload_file(zip_file, f.read()).execute_query()
    print(f"Đã upload {zip_file} lên SharePoint.")

    os.remove(zip_file)
    print(f"Đã xoá {zip_file} ở local.")

print("\nĐã xử lý xong tất cả các nhóm folder trong images.")
