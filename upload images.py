import os
import zipfile
import hashlib
import json
import time
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
MAX_ZIP_SIZE_MB = 200
MAX_ZIP_SIZE = MAX_ZIP_SIZE_MB * 1024 * 1024
tmp_download_folder = "__tmp_sharepoint__"
if not os.path.exists(tmp_download_folder):
    os.makedirs(tmp_download_folder)

# ==== Dọn file .txt: xoá .txt KHÔNG có 'comment' hoặc 'status' trong tên ====
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

# ==== Hàm nén ảnh gốc (inplace) ====
def compress_image_inplace(path, quality=70, max_size=(1920,1080)):
    try:
        img = Image.open(path)
        img.thumbnail(max_size, Image.LANCZOS)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.save(path, "JPEG", quality=quality, optimize=True)
    except Exception as e:
        print(f"Lỗi nén ảnh {path}: {e}")

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
            if ext not in ["jpg","jpeg","png"]:
                continue
            path = os.path.join(root, file)
            mtime = os.path.getmtime(path)
            size = os.path.getsize(path)
            key = os.path.relpath(path, folder)
            files_info[key] = [mtime, size]
            old = old_meta.get(key)
            if (old is None) or (old[0] != mtime or old[1] != size):
                compress_image_inplace(path, quality, max_size)
                nened = True
                mtime2 = os.path.getmtime(path)
                size2 = os.path.getsize(path)
                files_info[key] = [mtime2, size2]
    write_compress_meta(folder, files_info)
    return nened

# ==== Tiện ích chung ====
def md5sum(filename, bufsize=65536):
    h = hashlib.md5()
    with open(filename, "rb") as f:
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
    min_time = None
    for root, _, files in os.walk(folder):
        for file in files:
            path = os.path.join(root, file)
            ctime = os.path.getctime(path)
            if min_time is None or ctime < min_time:
                min_time = ctime
    if not min_time:
        return None
    return datetime.fromtimestamp(min_time).strftime("%Y%m")

def get_folder_size(folder):
    total = 0
    for dirpath, _, filenames in os.walk(folder):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.isfile(fp):
                total += os.path.getsize(fp)
    return total

# ==== Gộp theo ngày ====
IMAGE_EXTS = {".jpg",".jpeg",".png",".webp",".tif",".tiff",".bmp",".heic",".heif"}

def is_image_file(path): return os.path.splitext(path)[1].lower() in IMAGE_EXTS

def parse_date_from_name(name):
    import re
    s = name.lower()
    pats = [
        r'(?P<y>20\d{2})[-_\.]?(?P<m>0[1-9]|1[0-2])[-_\.]?(?P<d>0[1-9]|[12]\d|3[01])',
        r'(?P<d>0[1-9]|[12]\d|3[01])[-_\.](?P<m>0[1-9]|1[0-2])[-_\.](?P<y>20\d{2})',
    ]
    for pat in pats:
        m = re.search(pat, s)
        if m:
            try:
                y,mth,d = int(m.group("y")),int(m.group("m")),int(m.group("d"))
                return datetime(y,mth,d)
            except: continue
    return None

def get_exif_datetime(path):
    try:
        with Image.open(path) as im:
            exif = getattr(im,"getexif",lambda:None)()
            if not exif: return None
            from PIL import ExifTags
            keys = {k for k,v in ExifTags.TAGS.items() if v in ("DateTimeOriginal","DateTimeDigitized","DateTime")}
            for k in keys:
                if k in exif:
                    val = str(exif.get(k))
                    try:
                        return datetime.strptime(val.replace("-",
                            ":").replace(".",
                            ":"), "%Y:%m:%d %H:%M:%S")
                    except: pass
    except: pass
    return None

def best_guess_date(path):
    dt = get_exif_datetime(path)
    if dt: return dt
    dt = parse_date_from_name(os.path.basename(path))
    if dt: return dt
    return datetime.fromtimestamp(os.path.getmtime(path))

def ensure_dir(p): os.makedirs(p, exist_ok=True)

def ensure_unique_path(dst_path):
    if not os.path.exists(dst_path): return dst_path
    base,ext = os.path.splitext(dst_path)
    i=1
    while True:
        cand=f"{base}_{i}{ext}"
        if not os.path.exists(cand): return cand
        i+=1

def normalize_basename(name):
    import re
    s=re.sub(r"\s+","_",name)
    s=re.sub(r"[^a-zA-Z0-9._-]","_",s)
    return s.lower()

def copy_and_resize(src,dst,max_side=(1920,1080),quality=80,force_ext=".jpg"):
    with Image.open(src) as im:
        im=ImageOps.exif_transpose(im)
        fmt="JPEG" if force_ext.lower() in (".jpg",".jpeg") else "WEBP"
        if fmt=="JPEG": im=im.convert("RGB")
        w,h=im.size
        max_current=max(w,h)
        max_target=max(max_side)
        if max_current>max_target:
            scale=max_target/float(max_current)
            im=im.resize((int(w*scale),int(h*scale)),Image.LANCZOS)
        ensure_dir(os.path.dirname(dst))
        im.save(dst,format=fmt,quality=quality,optimize=True)

def build_day_tree(source_folders, month_label, max_side=(1920,1080), quality=80, force_ext=".jpg"):
    build_root=os.path.abspath(f"{month_label}_byday")
    if os.path.exists(build_root):
        import shutil; shutil.rmtree(build_root)
    os.makedirs(build_root,exist_ok=True)
    for folder in source_folders:
        for root,_,files in os.walk(folder):
            for fname in files:
                src=os.path.join(root,fname)
                if is_image_file(src):
                    dt=best_guess_date(src)
                    day=dt.strftime("%Y-%m-%d")
                    base=normalize_basename(os.path.splitext(fname)[0])
                    ext=force_ext
                    out_dir=os.path.join(build_root,day)
                    out_path=ensure_unique_path(os.path.join(out_dir,f"{dt.strftime('%Y%m%d_%H%M%S')}_{base}{ext}"))
                    copy_and_resize(src,out_path,max_side=max_side,quality=quality,force_ext=ext)
                else:
                    rel_dir=os.path.relpath(root,os.path.commonpath(source_folders))
                    out_dir=os.path.join(build_root,rel_dir)
                    ensure_dir(out_dir)
                    import shutil
                    shutil.copy2(src,os.path.join(out_dir,fname))
    return build_root

# ==== Đăng nhập SharePoint ====
ctx_auth = AuthenticationContext(site_url)
if not ctx_auth.acquire_token_for_user(username, password):
    raise Exception("Không kết nối được SharePoint!")
ctx = ClientContext(site_url, ctx_auth)

# ==== Duyệt từng folder ====
folders=[os.path.join(local_images,f) for f in os.listdir(local_images) if os.path.isdir(os.path.join(local_images,f))]
folders.sort()
folders_by_month={}
for folder in folders:
    folder_name=os.path.basename(folder)
    print(f"Clean .txt trong {folder_name} ...")
    clean_unlabeled_txt(folder)
    print(f"Nén folder {folder_name} ...")
    compress_folder_inplace_smart(folder)
    thang=folder_month(folder)
    if not thang: continue
    folders_by_month.setdefault(thang,[]).append(folder)

# ==== Gom theo tháng & zip (gộp theo ngày trước khi zip) ====
for thang,month_folders in sorted(folders_by_month.items()):
    cur_group=[]; cur_size=0; group_idx=1
    for folder in month_folders:
        fsize=get_folder_size(folder)
        if cur_group and cur_size+fsize>MAX_ZIP_SIZE:
            zip_file=f"images_{thang}_part{group_idx}.zip" if group_idx>1 else f"images_{thang}.zip"
            # gộp theo ngày
            byday_root=build_day_tree(cur_group,thang)
            with zipfile.ZipFile(zip_file,"w",zipfile.ZIP_DEFLATED) as zipf:
                for root,_,files in os.walk(byday_root):
                    for file in files:
                        full_path=os.path.join(root,file)
                        arcname=os.path.relpath(full_path,os.path.dirname(byday_root))
                        zipf.write(full_path,arcname)
            # check MD5 + upload
            local_md5=md5sum(zip_file)
            remote_zip_path=os.path.join(tmp_download_folder,zip_file)
            exists=download_file_from_sharepoint(ctx,upload_folder_sharepoint,zip_file,remote_zip_path)
            if exists:
                if md5sum(remote_zip_path)==local_md5:
                    print(f"{zip_file} không thay đổi.")
                    os.remove(zip_file); os.remove(remote_zip_path)
                else:
                    with open(zip_file,"rb") as fz:
                        ctx.web.get_folder_by_server_relative_url(upload_folder_sharepoint).upload_file(zip_file,fz.read()).execute_query()
                    print(f"Uploaded {zip_file}")
                    os.remove(zip_file); os.remove(remote_zip_path)
            else:
                with open(zip_file,"rb") as fz:
                    ctx.web.get_folder_by_server_relative_url(upload_folder_sharepoint).upload_file(zip_file,fz.read()).execute_query()
                print(f"Uploaded {zip_file}")
                os.remove(zip_file)
            cur_group=[]; cur_size=0; group_idx+=1
        cur_group.append(folder); cur_size+=fsize

    if cur_group:
        zip_file=f"images_{thang}_part{group_idx}.zip" if group_idx>1 else f"images_{thang}.zip"
        byday_root=build_day_tree(cur_group,thang)
        with zipfile.ZipFile(zip_file,"w",zipfile.ZIP_DEFLATED) as zipf:
            for root,_,files in os.walk(byday_root):
                for file in files:
                    full_path=os.path.join(root,file)
                    arcname=os.path.relpath(full_path,os.path.dirname(byday_root))
                    zipf.write(full_path,arcname)
        local_md5=md5sum(zip_file)
        remote_zip_path=os.path.join(tmp_download_folder,zip_file)
        exists=download_file_from_sharepoint(ctx,upload_folder_sharepoint,zip_file,remote_zip_path)
        if exists:
            if md5sum(remote_zip_path)==local_md5:
                print(f"{zip_file} không thay đổi.")
                os.remove(zip_file); os.remove(remote_zip_path)
            else:
                with open(zip_file,"rb") as fz:
                    ctx.web.get_folder_by_server_relative_url(upload_folder_sharepoint).upload_file(zip_file,fz.read()).execute_query()
                print(f"Uploaded {zip_file}")
                os.remove(zip_file); os.remove(remote_zip_path)
        else:
            with open(zip_file,"rb") as fz:
                ctx.web.get_folder_by_server_relative_url(upload_folder_sharepoint).upload_file(zip_file,fz.read()).execute_query()
            print(f"Uploaded {zip_file}")
            os.remove(zip_file)

print("\nHoàn tất: đã gộp theo ngày, nén ảnh bản sao và upload lên SharePoint.")
