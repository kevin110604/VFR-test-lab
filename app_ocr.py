# app_ocr_pro.py
import io, base64, os, sys, statistics, math
from typing import Tuple, Dict, Any, List, Optional
from flask import Flask, request, jsonify, render_template_string
from PIL import Image, ImageOps, UnidentifiedImageError
import pytesseract
import cv2
import numpy as np

# (Tuỳ chọn) Windows: chỉ định đường dẫn tesseract nếu PATH chưa có
# if sys.platform.startswith("win") and os.path.exists(r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
#     pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

app = Flask(__name__)

HTML = r"""
<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>OCR demo — Chụp/Upload ảnh → In toàn bộ text</title>
<style>
  :root { color-scheme: light dark; }
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; margin: 16px; line-height: 1.45; }
  .wrap { max-width: 980px; margin: 0 auto; display: grid; gap: 16px; }
  .row { display: grid; gap: 16px; grid-template-columns: 1fr 1fr; }
  @media (max-width: 900px){ .row { grid-template-columns: 1fr; } }
  .card { border: 1px solid #5553; border-radius: 12px; padding: 12px; }
  video, img, canvas { width: 100%; border: 1px solid #5553; border-radius: 10px; }
  textarea { width: 100%; min-height: 260px; border-radius: 10px; border: 1px solid #5555; padding: 10px; }
  button { padding: 10px 14px; border-radius: 10px; border: 1px solid #4445; background: #eee; cursor: pointer; }
  .row-btns { display: flex; gap: 8px; flex-wrap: wrap; }
  .muted { opacity: .75; font-size: 13px; }
</style>
</head>
<body>
<div class="wrap">
  <h2>📷 OCR demo — Chụp/Upload ảnh → In toàn bộ text</h2>

  <div class="row">
    <div class="card">
      <h3>1) Chụp ảnh bằng camera</h3>
      <video id="video" autoplay playsinline muted></video>
      <div class="row-btns">
        <button id="btnStart">Bật camera</button>
        <button id="btnSnap">Chụp ảnh</button>
        <button id="btnSend">Gửi ảnh để OCR</button>
      </div>
      <canvas id="canvas" width="1280" height="720" style="display:none"></canvas>
      <img id="preview" alt="Xem trước ảnh chụp" />
      <div class="muted">Camera chỉ hoạt động trên <b>http://localhost</b> hoặc <b>HTTPS</b>.</div>
    </div>

    <div class="card">
      <h3>2) Upload ảnh từ máy</h3>
      <form id="formUpload">
        <input type="file" name="image" id="file" accept="image/*" required />
        <div class="row-btns"><button type="submit">Upload & OCR</button></div>
      </form>
      <div class="muted">Không giới hạn loại ảnh; server sẽ cố gắng đọc.</div>
    </div>
  </div>

  <div class="card">
    <h3>Kết quả OCR (toàn bộ text)</h3>
    <textarea id="out" placeholder="Text OCR sẽ hiển thị ở đây..."></textarea>
    <div class="muted" id="meta"></div>
  </div>

  <div class="muted">
    Mẹo: nếu chỉ cần 1 nhóm ký tự (vd. mã/ số), thêm query <code>?wl=0-9A-Z/-.</code> vào URL để tăng độ chính xác.
  </div>
</div>

<script>
const $ = (q)=>document.querySelector(q);
const video = $('#video'), canvas = $('#canvas'), preview = $('#preview'), out = $('#out'), meta = $('#meta');

let stream = null;
$('#btnStart').addEventListener('click', async ()=>{
  try{
    stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false });
    video.srcObject = stream;
  }catch(e){ alert('Không truy cập được camera: ' + e); }
});

$('#btnSnap').addEventListener('click', ()=>{
  if(!stream){ alert('Bật camera trước.'); return; }
  const w = video.videoWidth || 1280, h = video.videoHeight || 720;
  canvas.width = w; canvas.height = h;
  canvas.getContext('2d').drawImage(video, 0, 0, w, h);
  preview.src = canvas.toDataURL('image/jpeg');
});

$('#btnSend').addEventListener('click', async ()=>{
  if(!preview.src){ alert('Hãy chụp ảnh trước.'); return; }
  const b64 = preview.src.split(',')[1];
  const r = await fetch('/ocr_base64' + location.search, {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({image_base64:b64})});
  const data = await r.json();
  out.value = data.error ? ('Lỗi: ' + data.error) : (data.text || '');
  meta.textContent = data.error ? '' : (`pipeline: ${data.pipeline}, psm: ${data.psm}, conf≈${data.confidence}`);
});

$('#formUpload').addEventListener('submit', async (e)=>{
  e.preventDefault();
  const fd = new FormData(e.target);
  const r = await fetch('/ocr' + location.search, { method:'POST', body: fd });
  const data = await r.json();
  out.value = data.error ? ('Lỗi: ' + data.error) : (data.text || '');
  meta.textContent = data.error ? '' : (`pipeline: ${data.pipeline}, psm: ${data.psm}, conf≈${data.confidence}`);
});
</script>
</body>
</html>
"""

# ----------------- UTIL & PREPROCESS -----------------

def _fix_orientation(pil: Image.Image) -> Image.Image:
    try:
        return ImageOps.exif_transpose(pil)
    except Exception:
        return pil

def _resize_for_ocr(img: np.ndarray, target_long_side: int = 2000) -> np.ndarray:
    h, w = img.shape[:2]
    long_side = max(h, w)
    if long_side >= target_long_side:
        return img
    scale = target_long_side / float(long_side)
    new_w, new_h = int(w * scale), int(h * scale)
    return cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_LINEAR)

def _deskew(gray: np.ndarray) -> np.ndarray:
    try:
        bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        coords = np.column_stack(np.where(bw > 0))
        if coords.size == 0: return gray
        angle = cv2.minAreaRect(coords)[-1]
        angle = -(90 + angle) if angle < -45 else -angle
        (h, w) = gray.shape[:2]
        M = cv2.getRotationMatrix2D((w//2, h//2), angle, 1.0)
        return cv2.warpAffine(gray, M, (w, h), flags=cv2.INTER_LINEAR, borderMode=cv2.BORDER_REPLICATE)
    except Exception:
        return gray

def _clahe(gray: np.ndarray) -> np.ndarray:
    return cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8)).apply(gray)

def _denoise(gray: np.ndarray) -> np.ndarray:
    # Bilateral giữ biên, nếu chậm có thể đổi sang medianBlur(3)
    return cv2.bilateralFilter(gray, d=7, sigmaColor=60, sigmaSpace=60)

def _sharpen(gray: np.ndarray) -> np.ndarray:
    blur = cv2.GaussianBlur(gray, (0,0), 1.0)
    sharp = cv2.addWeighted(gray, 1.6, blur, -0.6, 0)
    return np.clip(sharp, 0, 255).astype(np.uint8)

def _morph_open_close(bin_img: np.ndarray) -> np.ndarray:
    # mở rồi đóng để xóa nhiễu nhỏ và siết nét chữ
    k1 = cv2.getStructuringElement(cv2.MORPH_RECT, (2,2))
    k2 = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
    opened = cv2.morphologyEx(bin_img, cv2.MORPH_OPEN, k1, iterations=1)
    closed = cv2.morphologyEx(opened, cv2.MORPH_CLOSE, k2, iterations=1)
    return closed

def _pipelines(pil: Image.Image) -> List[Tuple[str, Image.Image]]:
    """Tạo nhiều biến thể ảnh để thử OCR; trả (tên, ảnh PIL)."""
    pil = _fix_orientation(pil.convert("RGB"))
    rgb = np.array(pil)
    rgb = _resize_for_ocr(rgb, target_long_side=2100)

    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    gray = _deskew(gray)
    gray = _clahe(gray)
    gray = _denoise(gray)
    gray = _sharpen(gray)

    # Otsu + morphology
    _, th_otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    th_otsu = _morph_open_close(th_otsu)

    # Adaptive (Gaussian)
    th_adp = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 33, 11)
    th_adp = _morph_open_close(th_adp)

    return [
        ("gray_clahe_sharp", Image.fromarray(gray)),
        ("binary_otsu_morph", Image.fromarray(th_otsu)),
        ("binary_adapt_morph", Image.fromarray(th_adp)),
    ]

# ------------- TESSERACT CONFIGS & OCR CORE -------------

BASE_CFG = "-c preserve_interword_spaces=1 -c user_defined_dpi=300"
TESS_CFGS = [
    ("6",  f"{BASE_CFG} --oem 3 --psm 6"),   # khối văn bản
    ("4",  f"{BASE_CFG} --oem 3 --psm 4"),   # nhiều cột
    ("11", f"{BASE_CFG} --oem 3 --psm 11"),  # thưa/rải rác
    ("3",  f"{BASE_CFG} --oem 3 --psm 3"),   # auto fully
]

def _ocr_with_conf(pil_img: Image.Image, lang: str, config: str) -> Tuple[str, float]:
    text = pytesseract.image_to_string(pil_img, lang=lang, config=config)
    # lấy conf trung bình từ image_to_data
    data = pytesseract.image_to_data(pil_img, lang=lang, config=config, output_type=pytesseract.Output.DICT)
    confs = []
    for c in data.get("conf", []):
        try:
            c = float(c)
            if c >= 0: confs.append(c)
        except Exception:
            pass
    avg = round(statistics.mean(confs), 1) if confs else 0.0
    return text, avg

def _best_ocr(pil: Image.Image, whitelist: Optional[str] = None) -> Dict[str, Any]:
    langs_to_try = ["vie+eng", "eng"]
    best = {"text":"", "confidence":0.0, "pipeline":"", "psm":""}

    # Thêm whitelist nếu có (lưu ý: chỉ dùng khi bạn chắc kiểu ký tự)
    wl_cfg = ""
    if whitelist:
        wl_cfg = f" -c tessedit_char_whitelist={whitelist} "

    for lang in langs_to_try:
        for pname, pimg in _pipelines(pil):
            for psm, cfg in TESS_CFGS:
                try:
                    text, conf = _ocr_with_conf(pimg, lang, cfg + wl_cfg)
                except pytesseract.TesseractError:
                    continue
                # score = conf (có thể cộng thêm len(text) để tránh conf cao nhưng text ngắn bất thường)
                score = conf + min(len(text)/300.0, 5)  # ưu tiên text dài hơn một chút
                if score > best["confidence"]:
                    best = {"text": text, "confidence": round(conf,1), "pipeline": pname, "psm": psm}
        if best["text"].strip():
            break
    return best

# -------------------------- ROUTES --------------------------

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/ocr", methods=["POST"])
def ocr_upload():
    wl = request.args.get("wl")  # tùy chọn whitelist, ví dụ wl=0-9A-Z/-. 
    if "image" not in request.files:
        return jsonify({"error":"Thiếu file ảnh (field name: image)."}), 400
    f = request.files["image"]
    if not f or f.filename == "":
        return jsonify({"error":"File ảnh trống."}), 400
    try:
        pil = Image.open(io.BytesIO(f.read()))
    except UnidentifiedImageError:
        return jsonify({"error":"File không phải ảnh hợp lệ."}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    result = _best_ocr(pil, whitelist=wl)
    return jsonify({"text": result["text"], "confidence": result["confidence"], "pipeline": result["pipeline"], "psm": result["psm"]})

@app.route("/ocr_base64", methods=["POST"])
def ocr_base64():
    wl = request.args.get("wl")
    data = request.get_json(silent=True) or {}
    b64 = data.get("image_base64")
    if not b64:
        return jsonify({"error":"Thiếu image_base64."}), 400
    try:
        pil = Image.open(io.BytesIO(base64.b64decode(b64)))
    except UnidentifiedImageError:
        return jsonify({"error":"Chuỗi base64 không phải ảnh hợp lệ."}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    result = _best_ocr(pil, whitelist=wl)
    return jsonify({"text": result["text"], "confidence": result["confidence"], "pipeline": result["pipeline"], "psm": result["psm"]})

if __name__ == "__main__":
    # Dev: localhost để test camera; khi public: Waitress/Caddy (HTTPS).
    app.run(host="0.0.0.0", port=8080, debug=False)
