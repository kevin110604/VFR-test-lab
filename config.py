import os

ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png', 'bmp'}
UPLOAD_FOLDER = "images"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
local_main = r"E:\VFR\ds sản phẩm test với qr.xlsx"
local_complete = r"E:\VFR\completed_items.xlsx"
qr_folder = r"E:\VFR\qr_labels"
PASSWORD = "123"
SECRET_KEY = "vfr_secret_key_123"
SO_GIO_TEST = 0.02
ALL_SLOTS = (
    [f"Ke1-A{i}" for i in range(1, 9)] + [f"Ke1-B{i}" for i in range(1, 9)] +
    [f"Ke2-A{i}" for i in range(1, 4)] + [f"Ke2-B{i}" for i in range(1, 4)] +
    [f"Ke3-A{i}" for i in range(1, 9)] + [f"Ke3-B{i}" for i in range(1, 9)]
)
SAMPLE_STORAGE = {}  # {slot_id: {'sample_name': ..., ...}}

TEST_GROUPS = [
    ("ban_us", "BÀN US"),
    ("ban_eu", "BÀN EU"),
    ("ghe_us", "GHẾ US"),
    ("ghe_eu", "GHẾ EU"),
    ("tu_us",  "TỦ US"),
    ("tu_eu",  "TỦ EU"),
    ("giuong", "GIƯỜNG"),
    ("guong",  "GƯƠNG"),
    ("testkhac", "TEST KHÁC"),
]