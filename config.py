import os

ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png', 'bmp'}
UPLOAD_FOLDER = "images"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
local_main = "ds sản phẩm test với qr.xlsx"
local_complete = "completed_items.xlsx"
qr_folder = "qr_labels"
TEAMS_WEBHOOK_URL = "https://mphcmiuedu.webhook.office.com/webhookb2/11e4d9d9-a3bf-4f77-9947-4c790d2c90e0@a7380202-eb54-415a-9b66-4d9806cfab42/IncomingWebhook/69ca9bf67d3342c1984109a9f3073faf/81e98e9d-7a3f-492d-9f99-3ac5a5ecbbf3/V2ou84tjRa9HGhGQHzzzaVkzig_fT-ACqaRCaq11ertlo1"
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