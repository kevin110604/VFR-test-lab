import os
import unicodedata
import re
from config import UPLOAD_FOLDER, ALLOWED_EXTENSIONS

def safe_filename(filename):
    """Chuyển filename sang dạng an toàn, chỉ giữ lại ký tự chữ-số-gạch-dưới và giữ lại đuôi file."""
    filename = unicodedata.normalize('NFKD', filename).encode('ascii', 'ignore').decode('ascii')
    filename = re.sub(r'[^a-zA-Z0-9_.-]', '_', filename)
    return filename

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_img_urls(report, tag=None, upload_folder="images"):
    urls = []
    folder = os.path.join(upload_folder, report)
    if not os.path.exists(folder): return urls
    for fname in os.listdir(folder):
        if tag and not fname.startswith(f"{tag}_"):
            continue
        if fname.rsplit('.', 1)[-1].lower() in ALLOWED_EXTENSIONS:
            urls.append(f"/images/{report}/{fname}")
    return sorted(urls)