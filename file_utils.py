import json
import os
import openpyxl

# =========================
# Internal helpers
# =========================

def _atomic_write(path: str, text: str) -> None:
    """
    Ghi file theo kiểu atomic:
    - Ghi ra file tạm .tmp
    - flush + fsync
    - os.replace để thay thế (atomic trên hầu hết OS)
    """
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(text)
        try:
            f.flush()
            os.fsync(f.fileno())
        except Exception:
            # Một số FS/OS có thể không hỗ trợ fsync hoặc không cần thiết
            pass
    os.replace(tmp, path)


def _backup_path(path: str) -> str:
    """
    Trả về đường dẫn backup cạnh file gốc:
    - x.json -> x_backup.json
    - x     -> x_backup
    """
    base, ext = os.path.splitext(path)
    return f"{base}_backup{ext}" if ext else (path + "_backup")


# =========================
# JSON helpers
# =========================

def safe_read_json(path, default=None):
    """
    Đọc JSON an toàn.
    - Nếu file không tồn tại hoặc parse lỗi -> trả về default (mặc định []).
    - Không raise để không chặn luồng nghiệp vụ.
    """
    if not os.path.exists(path):
        return default if default is not None else []
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default if default is not None else []


def safe_write_json(path, data):
    """
    Ghi JSON an toàn (atomic) vào FILE CHÍNH.
    - KHÔNG ghi/đụng chạm tới file backup.
    - Dùng khi cập nhật trạng thái hiện tại (thêm/sửa/xoá) của hệ thống.
    """
    folder = os.path.dirname(path) or "."
    os.makedirs(folder, exist_ok=True)
    _atomic_write(path, json.dumps(data, ensure_ascii=False, indent=2))


def safe_append_backup_json(path, new_records):
    """
    Append-only vào FILE BACKUP (…_backup.json) cạnh file chính.
    - Không bao giờ xoá/sửa dữ liệu backup; chỉ nối thêm bản ghi mới.
    - 'new_records' có thể là 1 dict hoặc list[dict].
    - Nếu backup chưa tồn tại -> tạo mới dạng list.
    """
    bak = _backup_path(path)
    os.makedirs(os.path.dirname(bak) or ".", exist_ok=True)

    cur = safe_read_json(bak, default=[])
    if not isinstance(cur, list):
        # Nếu file backup hiện tại không phải list (bị thay đổi từ trước),
        # vẫn đảm bảo an toàn bằng cách chuyển về list.
        cur = []

    if isinstance(new_records, list):
        cur.extend(new_records)
    else:
        cur.append(new_records)

    _atomic_write(bak, json.dumps(cur, ensure_ascii=False, indent=2))


# =========================
# Text helpers
# =========================

def safe_read_text(path, default=""):
    """
    Đọc text an toàn. Nếu không tồn tại -> trả về default.
    """
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def safe_write_text(path, text):
    """
    Ghi text an toàn (atomic).
    - KHÔNG tạo backup để tránh phát sinh file ngoài ý muốn.
    """
    folder = os.path.dirname(path) or "."
    os.makedirs(folder, exist_ok=True)
    _atomic_write(path, text)


# =========================
# Excel helpers
# =========================

def safe_load_excel(path):
    """
    Mở workbook Excel (openpyxl). Hàm này giữ nguyên hành vi mặc định.
    """
    return openpyxl.load_workbook(path)


def safe_save_excel(wb, path):
    """
    Lưu workbook Excel. Giữ nguyên hành vi mặc định (không backup) để tránh
    xung đột với openpyxl khi replace file đang mở bởi tiến trình khác.
    """
    wb.save(path)
