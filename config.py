import os

ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png', 'bmp'}
UPLOAD_FOLDER = "images"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
local_main = "ds san pham test voi qr.xlsx"
local_complete = "completed_items.xlsx"
qr_folder = "qr_labels"
TEAMS_WEBHOOK_URL_TRF = "https://jonathancharles.webhook.office.com/webhookb2/c1820b56-e3c0-4435-b1a6-1630c0f8da85@064944f6-1e04-4050-b3e1-e361758625ec/IncomingWebhook/09c95cb9a4d5487c8b8aac5ef36f6d1d/169de55a-0196-4de6-b160-7a456bce2292/V2me37ZNeF3_Z1CnaOm4naD_tJ0TptNjr_rJRwga6qBSg1"
TEAMS_WEBHOOK_URL_RATE = "https://jonathancharles.webhook.office.com/webhookb2/c1820b56-e3c0-4435-b1a6-1630c0f8da85@064944f6-1e04-4050-b3e1-e361758625ec/IncomingWebhook/b4e8f13a441b4d7cbd7470253710b107/169de55a-0196-4de6-b160-7a456bce2292/V2oi3fuDogcw1BlJSoJgFc5XVn4xPpWv3eZKCAFjThbkI1"
TEAMS_WEBHOOK_URL_COUNT = "https://jonathancharles.webhook.office.com/webhookb2/c1820b56-e3c0-4435-b1a6-1630c0f8da85@064944f6-1e04-4050-b3e1-e361758625ec/IncomingWebhook/b344a6ce0fbe4f60bbfe8b16ba0d203c/169de55a-0196-4de6-b160-7a456bce2292/V2pE_UIJyO6XQrj7OFUSvU8xKKGS8IzOdwmz8LjLsWT501"
PASSWORD_STL = "stl"
PASSWORD_WTL = "wtl"
PASSWORD_VFR3 = "vfr3"
MASTER_PASSWORD = os.getenv("MASTER_PASSWORD", "admin")
SECRET_KEY = "vfr_secret_key_123"
SO_GIO_TEST = 50
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
# >>> ADD: ánh xạ loại test -> file template .docx (đặt cùng thư mục app.py)
TEMPLATE_MAP = {
    # Structure
    'bed': 'BED - GLB.docx',
    'chair_us': 'CHAIR - US.docx',
    'chair_uk': 'CHAIR - EU.docx',
    'table_us': 'TABLE - US.docx',
    'table_uk': 'TABLE - EU,UK.docx',
    'cabinet_us': 'STORAGE - US.docx',
    'cabinet_uk': 'STORAGE - EU.docx',
    'mirror': 'MIRROR.docx',
    'other': 'OTHER.docx',

    # Material – Finishing
    'material_indoor_chuyen': 'MAT-IN-CHUYEN.docx',
    'material_indoor_qa': 'MAT-IN-QA.docx',
    'material_indoor_stone': 'MAT-IN-STONE.docx',
    'material_indoor_metal': 'MAT-IN-METAL.docx',
    'material_outdoor': 'MAT-OUT.docx',
    'line_test': 'LINE-TEST.docx',
    'hot_cold_test': 'HOT-COLD.docx',

    # Transit
    'transit_2c_np': 'TRANSIT-2C-NP.docx',
    'transit_rh_np': 'TRANSIT-RH-NP.docx',
    'transit_181_lt68': 'TRANSIT-181-LT68.docx',
    'transit_3a': 'TRANSIT-3A.docx',
    'transit_3b_np': 'TRANSIT-3B-NP.docx',

    'transit_2c_pallet': 'TRANSIT-2C-PL.docx',
    'transit_rh_pallet': 'TRANSIT-RH-PL.docx',
    'transit_181_gt68': 'TRANSIT-181-GT68.docx',
    'transit_3b_pallet': 'TRANSIT-3B-PL.docx',
}
