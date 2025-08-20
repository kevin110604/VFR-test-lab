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