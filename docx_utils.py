# ========================= docx_utils.py (FULL — TRF giữ nguyên, BED sửa tick) =========================
import os
import re
import time
import uuid
import tempfile
from io import BytesIO
from openpyxl import load_workbook
from docx import Document

from config import local_main  # local_main: đường dẫn/tệp Excel gốc
from excel_utils import _find_report_col

# Fuzzy-match Excel <-> label form
import unicodedata
try:
    import pandas as pd
except Exception:
    pd = None

# === CONFIG ===
WORD_TEMPLATE = "FORM-QAD-011-TEST REQUEST FORM (TRF).docx"
PDF_OUTPUT_FOLDER = os.path.join("static", "TFR")

# =========================
# Blank-detector for Excel
# =========================
BLANK_TOKENS = {"", "-", "—", "–"}

def _normalize_to_check_blank(v):
    if v is None:
        return True, ""
    if isinstance(v, str):
        s = (
            v.replace("\u00A0", "")
             .replace("\u200B", "")
             .replace("\r", "")
             .replace("\n", "")
             .replace("\t", "")
             .strip()
        )
        return (s in BLANK_TOKENS), s
    return False, str(v)

def get_first_empty_report_all_blank(excel_path):
    """Tìm mã report đầu tiên có toàn bộ cột C..X đều trống (dựa file Excel tại local_main)."""
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active
    report_col = _find_report_col(ws)  # <- TÌM CỘT ĐỘNG

    for row in range(2, ws.max_row + 1):
        all_mid_empty = True
        for col in range(3, 25):  # C..X
            is_blank, _ = _normalize_to_check_blank(ws.cell(row=row, column=col).value)
            if not is_blank:
                all_mid_empty = False
                break
        if all_mid_empty:
            report_no = ws.cell(row=row, column=report_col).value
            if report_no is not None and str(report_no).strip():
                wb.close()
                return str(report_no).strip()
    wb.close()
    return None

# =========================
# Checkbox mapping cho DOCX (TRF)
# =========================
def build_label_value_map(data):
    """
    Sinh map {label_in_template: bool} cho các checkbox trong file Word (TRF).
    - Nới lỏng 'test_group': khớp cả “CONSTRUCTION” và “CONSTRUCTION TEST”.
    - Giữ logic '...TH' cho test_status (nth).
    - Sinh cờ N/A cho các trường văn bản.
    """
    label_groups = {
        "test_group": [
            "MATERIAL TEST",
            "FINISHING TEST",
            "CONSTRUCTION TEST",
            "TRANSIT TEST",
            "ENVIRONMENTAL TEST",
        ],
        "test_status": ["1ST", "2ND", "3RD", "...TH"],
        "furniture_testing": ["INDOOR", "OUTDOOR"],
        "sample_return": ["YES", "NO"],
    }

    def _eq_relaxed(label: str, value: str, group: str) -> bool:
        L = (label or "").strip().upper()
        V = (value or "").strip().upper()
        if not V:
            return False
        if group == "test_group":
            if V.endswith(" TEST"): V2 = V[:-5].strip()
            else: V2 = V
            if L.endswith(" TEST"): L2 = L[:-5].strip()
            else: L2 = L
            return (V == L) or (V2 == L) or (V == L2) or (V2 == L2)
        if group == "test_status" and label == "...TH":
            return V.endswith("TH") and V not in {"1ST", "2ND", "3RD"}
        return V == L

    label_value_map = {}
    for group, labels in label_groups.items():
        value = data.get(group, None)
        if value is None or (isinstance(value, str) and not value.strip()):
            for label in labels:
                label_value_map[label] = False
        else:
            for label in labels:
                label_value_map[label] = _eq_relaxed(label, str(value), group)

    # N/A flags
    for field in ["sample_description", "item_code", "supplier", "subcon"]:
        val = str(data.get(field, "")).strip().upper()
        label_value_map[f"{field.upper()} N/A"] = (val == "N/A")

    return label_value_map

# =========================
# Tick checkbox theo ĐÚNG NHÃN (TRF) – regex trên text, thay luôn p.text / cell.text
# =========================
def _label_regex(label: str) -> re.Pattern:
    """Regex khớp '☐/☑' ngay TRƯỚC nhãn (khớp lỏng khoảng trắng/tab)."""
    cleaned = re.sub(r'[_\.\-]+', ' ', (label or '').strip())
    parts = [p for p in cleaned.split() if p]
    pattern = r'(☐|☑)\s*' + r'\s*'.join(re.escape(p) for p in parts)
    return re.compile(pattern, flags=re.IGNORECASE)

def tick_unicode_checkbox_by_label(doc: Document, label_value_map):
    """Duyệt paragraph + table cell, chỉ thay ký tự checkbox nằm TRƯỚC nhãn khớp."""
    compiled = [(_label_regex(label), bool(value)) for label, value in label_value_map.items()]

    def toggle_text(txt: str) -> str:
        if not txt or ('☐' not in txt and '☑' not in txt):
            return txt
        for pat, value in compiled:
            def _repl(m):
                return ('☑' if value else '☐') + m.group(0)[1:]
            txt = pat.sub(_repl, txt)
        return txt

    for p in doc.paragraphs:
        t = p.text
        if ("☐" in t) or ("☑" in t):
            p.text = toggle_text(t)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text
                if ("☐" in t) or ("☑" in t):
                    cell.text = toggle_text(t)

# =========================
# Convert PDF (best-effort)
# =========================
def try_convert_to_pdf(docx_path, pdf_path):
    try:
        import pythoncom
        pythoncom.CoInitialize()
        from docx2pdf import convert
        convert(docx_path, pdf_path)
    except Exception as e:
        import traceback
        print("Không thể convert PDF:", e)
        traceback.print_exc()

# =========================
# Per-report lock + atomic save
# =========================
def _lock_path_for(report_no: str) -> str:
    return os.path.join(tempfile.gettempdir(), f"tfr_{report_no}.lock")

def _acquire_lock(path: str, timeout=30):
    t0 = time.time()
    fd = None
    while True:
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
            return fd
        except FileExistsError:
            if time.time() - t0 > timeout:
                try: os.unlink(path)
                except: pass
            time.sleep(0.05)
        except Exception:
            time.sleep(0.05)

def _release_lock(fd, path: str):
    try: os.close(fd)
    except: pass
    try: os.unlink(path)
    except: pass

def _atomic_save_docx(doc: Document, out_path: str):
    tmp = f"{out_path}.tmp-{uuid.uuid4().hex}"
    doc.save(tmp)
    os.replace(tmp, out_path)  # atomic trên cùng filesystem

# =========================
# Fill DOCX + xuất PDF (TRF)
# =========================
def fill_docx_and_export_pdf(data, fixed_report_no=None):
    """
    - Nếu fixed_report_no có giá trị => dùng y nguyên.
    - Ngược lại => lấy report_no trống đầu tiên theo C..X đều trống.
    - Điền vào template, tick checkbox, atomic save DOCX, cố gắng convert PDF.
    """
    if fixed_report_no and str(fixed_report_no).strip():
        report_no = str(fixed_report_no).strip()
    else:
        report_no = get_first_empty_report_all_blank(local_main)
        if not report_no:
            raise Exception("Không còn mã report trống trong Excel.")

    data = dict(data)
    data["report_no"] = report_no

    doc = Document(WORD_TEMPLATE)

    mapping = {
        "requestor": "requestor",
        "department": "department",
        "requested date": "request_date",
        "lab test report no.": "report_no",
        "sample description": "sample_description",
        "item code": "item_code",
        "quantity": "quantity",
        "supplier": "supplier",
        "subcon": "subcon",
        "test group": "test_group",
        "test status": "test_status",
        "furniture testing": "furniture_testing",
        "estimated completed date": "estimated_completion_date",
    }
    remark = data.get("remark", "")
    remark_written = False

    # Điền các ô theo mapping
    for table in doc.tables:
        nrows = len(table.rows)
        ncols = len(table.columns)
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                label = (
                    cell.text.strip().lower()
                    .replace("(mã item)", "")
                    .replace("(mã material)", "")
                    .replace("*", "")
                )
                # remark
                if not remark_written and ("other tests/instructions" in label or "remark" in label) and remark:
                    if i + 1 < nrows:
                        below_cell = table.rows[i + 1].cells[j]
                        if not (below_cell.text or "").strip():
                            below_cell.text = str(remark)
                            remark_written = True
                            continue
                # employee id / msnv
                if ("emp id" in label or "msnv" in label) and data.get("employee_id", ""):
                    if j + 1 < ncols:
                        target_cell = row.cells[j + 1]
                        if not (target_cell.text or "").strip():
                            target_cell.text = str(data["employee_id"])
                            continue
                # mapping fields
                for map_label, key in mapping.items():
                    if map_label in ["remark", "employee id"]:
                        continue
                    if map_label in label and key in data and str(data[key]).strip() != "":
                        if j + 1 < ncols:
                            target_cell = row.cells[j + 1]
                            # cho phép set "lab test report no." để ghi đè
                            if (target_cell.text or "").strip() == "" or "lab test report no." in label:
                                target_cell.text = str(data[key])

    # Tick checkbox (TRF)
    label_value_map = build_label_value_map(data)
    tick_unicode_checkbox_by_label(doc, label_value_map)

    # Output
    if not os.path.exists(PDF_OUTPUT_FOLDER):
        os.makedirs(PDF_OUTPUT_FOLDER)
    output_docx = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.docx")
    output_pdf  = os.path.join(PDF_OUTPUT_FOLDER, f"{report_no}.pdf")

    lock_path = _lock_path_for(report_no)
    fd = _acquire_lock(lock_path, timeout=30)
    try:
        _atomic_save_docx(doc, output_docx)
    finally:
        _release_lock(fd, lock_path)

    try_convert_to_pdf(output_docx, output_pdf)

    return output_docx, output_pdf, report_no

# =========================
# API cho app.py (TRF)
# =========================
def approve_request_fill_docx_pdf(req):
    """
    Hàm wrapper để app.allocate_unique_report_no() gọi:
      - Nếu req chứa 'report_no' -> điền đúng số này (validate ở bên app).
      - Nếu không có -> tự tìm dòng C..X trống để lấy report_no.
    Trả về: (pdf_path_hoặc_docx_path, report_no)
    """
    fixed = (req.get("report_no") or "").strip()
    # Chuẩn hóa ETD
    if "etd" in req và not req.get("estimated_completion_date"):
        req = dict(req)
        req["estimated_completion_date"] = req.get("etd")

    if fixed:
        out_docx, out_pdf, report_no = fill_docx_and_export_pdf(req, fixed_report_no=fixed)
    else:
        out_docx, out_pdf, report_no = fill_docx_and_export_pdf(req, fixed_report_no=None)

    if os.path.exists(out_pdf):
        return out_pdf, report_no
    return out_docx, report_no


# =====================================================================
# ==================  BED COVER AUTO-FILL (Fuzzy)  ====================
# =====================================================================

# ---- Chuẩn hoá & chấm điểm giống/na ná ----
def _bed_norm(s: str) -> str:
    """Chuẩn hoá mạnh: lower + bỏ dấu TV + thay mọi ký tự ngoài [a-z0-9] bằng space (loại cả '#')."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", str(s).strip().lower())
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _bed_tokens(s: str) -> set:
    return set(_bed_norm(s).split()) if s else set()

def _bed_overlap_score(label: str, colname: str) -> float:
    ln = _bed_norm(label)
    cn = _bed_norm(colname)
    if not ln or not cn:
        return 0.0
    score = 0.0
    if ln == cn: score += 5.0
    if ln in cn: score += 3.0
    if cn in ln: score += 1.0
    lt = _bed_tokens(label); ct = _bed_tokens(colname)
    if lt and ct:
        inter = lt & ct
        score += 4.0 * (len(inter)/max(1,len(lt))) + 2.0 * (len(inter)/max(1,len(ct)))
    return score

def _bed_pick_best_column(excel_columns, label_text: str, preferred_aliases=None) -> str:
    """
    Chọn cột tốt nhất + bonus alias; **tránh 'QR Code'** khi label là Item code.
    """
    preferred_aliases = preferred_aliases or []
    is_item_code = _bed_norm(label_text) in {"item material code", "item code", "item material"}
    blacklist = {"qr code"} if is_item_code else set()

    best_col, best_score = "", -1.0
    for c in excel_columns:
        if _bed_norm(c) in blacklist:
            continue
        sc = _bed_overlap_score(label_text, c)
        if c in preferred_aliases: sc += 2.0
        if sc > best_score:
            best_col, best_score = c, sc
    return best_col if best_score > 0.8 else ""

def _bed_val(row, colname: str) -> str:
    if not colname:
        return ""
    try:
        v = row.get(colname, "")
    except Exception:
        v = ""
    if pd is not None and hasattr(pd, "isna") and pd.isna(v):
        return ""
    return str(v).strip()

def _smart_excel_path(path_or_name: str) -> str:
    """
    Nhận vào:
      - Đường dẫn file Excel đầy đủ, hoặc
      - Tên file Excel, hoặc
      - Rỗng/None -> dùng thẳng local_main
    Trả về đường dẫn file Excel hợp lệ, ưu tiên local_main nếu đầu vào không tồn tại.
    """
    if not path_or_name:
        return local_main
    p = str(path_or_name)
    if os.path.exists(p) and os.path.isfile(p):
        return p
    try:
        if os.path.isdir(local_main):
            cand = os.path.join(local_main, p)
            if os.path.exists(cand):
                return cand
    except Exception:
        pass
    return local_main

def _bed_load_excel_df(excel_path_or_name: str):
    if pd is None:
        raise RuntimeError("pandas chưa cài (pd is None)")
    excel_path = _smart_excel_path(excel_path_or_name)
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel không tồn tại: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _is_placeholder_dash(txt: str) -> bool:
    """Ô value coi là placeholder nếu rỗng hoặc chỉ chứa '-', '—', '–' (kể cả khoảng trắng/xuống dòng)."""
    if txt is None:
        return True
    t = re.sub(r"[\s\r\n\t]+", "", str(txt))
    return t in BLANK_TOKENS

# ---- Style helpers ----
def _clone_first_run_style(src_cell):
    """Lấy font name & size từ run đầu tiên trong ô nhãn (nếu có)."""
    try:
        p = src_cell.paragraphs[0]
        if not p.runs:
            return None, None
        r = p.runs[0]
        return r.font.name, r.font.size
    except Exception:
        return None, None

def _set_cell_text_with_style(dst_cell, src_label_cell, text):
    """Ghi text vào ô value và áp font theo nhãn (nếu bắt được)."""
    dst_cell.text = ""
    p = dst_cell.paragraphs[0]
    run = p.add_run("" if text is None else str(text))
    fname, fsize = _clone_first_run_style(src_label_cell)
    if fname: run.font.name = fname
    if fsize: run.font.size = fsize

# ---------- RESULT inline ----------
def _set_result_inline_in_paragraph(paragraph, rating_text: str) -> bool:
    """Thay '-' sau 'RESULT:' ngay trong run, giữ style run chứa '-'."""
    if not rating_text or not paragraph.runs:
        return False
    full = "".join(r.text for r in paragraph.runs)
    m = re.search(r"(?i)RESULT\s*:\s*", full)
    if not m:
        return False
    start_after_colon = m.end()

    idx = 0
    run_index = None
    offset_in_run = None
    for ridx, r in enumerate(paragraph.runs):
        txt = r.text
        nxt = idx + len(txt)
        if nxt > start_after_colon and idx <= start_after_colon:
            offset_in_run = start_after_colon - idx
            tail = txt[offset_in_run:]
            m2 = re.match(r"[\s]*[–—-]+", tail)
            if m2:
                run_index = ridx
            break
        idx = nxt

    if run_index is None:
        for ridx in range(len(paragraph.runs)):
            if re.fullmatch(r"\s*[–—-]+\s*", paragraph.runs[ridx].text or ""):
                run_index = ridx
                offset_in_run = 0
                break
    if run_index is None:
        return False

    r = paragraph.runs[run_index]
    txt = r.text or ""
    head = txt[:offset_in_run] if offset_in_run else ""
    new_txt = re.sub(r"^\s*[–—-]+\s*", " " + str(rating_text),
                     (txt[offset_in_run:] if offset_in_run is not None else txt))
    r.text = head + new_txt
    return True

def _set_result_value(doc: Document, rating_text: str):
    """Tìm 'RESULT: -' và thay '-' bằng rating_text (ưu tiên ô bảng, fallback inline)."""
    if not rating_text:
        return False
    # bảng
    for t in doc.tables:
        for r in t.rows:
            cells = r.cells
            for j in range(len(cells) - 1):
                if (cells[j].text or "").strip().upper().startswith("RESULT:"):
                    cur = (cells[j+1].text or "").strip()
                    if cur in BLANK_TOKENS:
                        _set_cell_text_with_style(cells[j+1], cells[j], rating_text)
                        return True
    # inline
    for p in doc.paragraphs:
        if "RESULT" in (p.text or "") or "Result" in (p.text or ""):
            if _set_result_inline_in_paragraph(p, rating_text):
                return True
    return False

# ---------- Test time helpers (BED) ----------
# 1) Regex kiểu TRF nhưng "linh hoạt mũ" cho 1ˢᵗ/2ⁿᵈ/3ʳᵈ/4ᵗʰ
_SUP = {
    "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
    "s": "ˢ", "t": "ᵗ", "n": "ⁿ", "d": "ᵈ", "r": "ʳ", "h": "ʰ"
}
def _flex_token(tok: str) -> str:
    # biến mỗi ký tự thành (thường|mũ), ghép lại
    parts = []
    for ch in tok:
        low = ch.lower()
        sup = _SUP.get(low, "")
        if sup:
            parts.append(f"(?:{re.escape(ch)}|{re.escape(sup)})")
        else:
            parts.append(re.escape(ch))
    return "".join(parts)

def _bed_label_regex(label: str) -> re.Pattern:
    """
    Ví dụ label '1st test' -> pattern '(☐|☑)\s*1(?:s|ˢ)(?:t|ᵗ)\s*test'
    """
    cleaned = re.sub(r'[_\.\-]+', ' ', (label or '').strip().lower())
    tokens = [t for t in cleaned.split() if t]
    flex_tokens = [_flex_token(t) for t in tokens]
    pattern = r'(☐|☑)\s*' + r'\s*'.join(flex_tokens)
    return re.compile(pattern, flags=re.IGNORECASE)

def _tick_testtime_by_regex(doc: Document, picked_label: str) -> bool:
    """
    Dựa theo cơ chế tick của TRF: thay trực tiếp trên text.
    Đồng thời reset 3 ô còn lại.
    """
    labels = ["1st test", "2nd test", "3rd test", "4th test"]
    compiled = [(_bed_label_regex(lbl), lbl == picked_label) for lbl in labels]

    def toggle_text(txt: str) -> (str, bool):
        changed = False
        if not txt or ('☐' not in txt and '☑' not in txt):
            return txt, changed
        for pat, value in compiled:
            def _repl(m):
                nonlocal changed
                changed = True
                return ('☑' if value else '☐') + m.group(0)[1:]
            txt = pat.sub(_repl, txt)
        return txt, changed

    any_change = False
    for p in doc.paragraphs:
        new_text, ch = toggle_text(p.text)
        if ch:
            p.text = new_text
            any_change = True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                new_text, ch = toggle_text(cell.text)
                if ch:
                    cell.text = new_text
                    any_change = True
    return any_change

# 2) Fallback: nếu regex không bắt được (dấu ☒ nằm ô riêng), tìm đúng hàng và tick ở ô/trái
def _norm_ascii(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    supers_map = str.maketrans({
        "⁰":"0","¹":"1","²":"2","³":"3","⁴":"4","⁵":"5","⁶":"6","⁷":"7","⁸":"8","⁹":"9",
        "ᵃ":"a","ᵇ":"b","ᶜ":"c","ᵈ":"d","ᵉ":"e","ᶠ":"f","ᵍ":"g","ʰ":"h","ᶦ":"i","ʲ":"j","ᵏ":"k",
        "ˡ":"l","ᵐ":"m","ⁿ":"n","ᵒ":"o","ᵖ":"p","ʳ":"r","ˢ":"s","ᵗ":"t","ᵘ":"u","ᵛ":"v","ʷ":"w",
        "ˣ":"x","ʸ":"y","ᶻ":"z"
    })
    s = s.translate(supers_map)
    s = s.replace("^", "")
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def _canon_test_label(s: str) -> str:
    x = _norm_ascii(s)
    x = re.sub(r"\btest\s*(\d)\b", r"\1", x)
    x = x.replace("first", "1st").replace("second", "2nd").replace("third", "3rd").replace("fourth", "4th")
    x = re.sub(r"\b1\s*st\b", "1st", x)
    x = re.sub(r"\b2\s*nd\b", "2nd", x)
    x = re.sub(r"\b3\s*rd\b", "3rd", x)
    x = re.sub(r"\b4\s*th\b", "4th", x)
    if "1st" in x: return "1st test"
    if "2nd" in x: return "2nd test"
    if "3rd" in x: return "3rd test"
    if "4th" in x: return "4th test"
    return ""

def _para_has_checkbox(p) -> bool:
    for r in p.runs:
        if ("☐" in (r.text or "")) or ("☑" in (r.text or "")):
            return True
    t = p.text or ""
    return ("☐" in t) or ("☑" in t)

def _replace_first_checkbox_in_paragraph(p, tick=True) -> bool:
    src = "☐" if tick else "☑"
    dst = "☑" if tick else "☐"
    for r in p.runs:
        if src in (r.text or ""):
            r.text = r.text.replace(src, dst, 1)
            return True
    if src in (p.text or ""):
        p.text = (p.text or "").replace(src, dst, 1)
        return True
    return False

def _ensure_checkbox_in_paragraph(p):
    if _para_has_checkbox(p):
        return True
    if not p.runs:
        p.add_run("")
    p.runs[0].text = "☐ " + (p.runs[0].text or "")
    return True

def _norm_label_text(s: str) -> str:
    s = _norm_ascii(s)
    s = re.sub(r"\b1\s*st\s*test\b", "1st test", s)
    s = re.sub(r"\b2\s*nd\s*test\b", "2nd test", s)
    s = re.sub(r"\b3\s*rd\s*test\b", "3rd test", s)
    s = re.sub(r"\b4\s*th\s*test\b", "4th test", s)
    return s

def _row_is_testlabel_row(row) -> bool:
    for c in row.cells:
        n = _norm_label_text(c.text)
        if any(k in n for k in ["1st test", "2nd test", "3rd test", "4th test"]):
            return True
    return False

def _find_test_time_table(doc: Document):
    candidates = []
    for t in doc.tables:
        has_header = False
        hits = 0
        for r in t.rows:
            for c in r.cells:
                n = _norm_ascii(c.text)
                if "test time" in n or "test phrase" in n:
                    has_header = True
                if any(lbl in n for lbl in ["1st test","2nd test","3rd test","4th test"]):
                    hits += 1
        if hits:
            candidates.append((has_header, hits, t))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (not x[0], -x[1]))
    return candidates[0][2]

def _clear_ticks_only_in_test_rows(tbl):
    for r in tbl.rows:
        if not _row_is_testlabel_row(r):
            continue
        for c in r.cells:
            for p in c.paragraphs:
                changed = True
                while changed:
                    changed = _replace_first_checkbox_in_paragraph(p, tick=False)

def _locate_target_tick_position(row, want_label: str):
    for j, cell in enumerate(row.cells):
        for p in cell.paragraphs:
            if want_label in _norm_label_text(p.text):
                _ensure_checkbox_in_paragraph(p)
                return p
    label_idx = -1
    for j, cell in enumerate(row.cells):
        if want_label in _norm_label_text(cell.text):
            label_idx = j
            break
    if label_idx == -1:
        return None
    for j_left in range(label_idx - 1, -1, -1):
        left_cell = row.cells[j_left]
        if not left_cell.paragraphs:
            p = left_cell.add_paragraph("")
            _ensure_checkbox_in_paragraph(p)
            return p
        for p in left_cell.paragraphs:
            _ensure_checkbox_in_paragraph(p)
            return p
    return None

def _tick_testtime_checkbox_fallback(doc: Document, picked_label: str) -> bool:
    if not picked_label:
        return False
    tbl = _find_test_time_table(doc)
    if tbl is None:
        return False
    want = _norm_label_text(picked_label)
    target_para = None
    for r in tbl.rows:
        if not _row_is_testlabel_row(r):
            continue
        p = _locate_target_tick_position(r, want)
        if p is not None:
            target_para = p
            break
    if target_para is None:
        return False
    _clear_ticks_only_in_test_rows(tbl)
    _ensure_checkbox_in_paragraph(target_para)
    return _replace_first_checkbox_in_paragraph(target_para, tick=True)

# ---------- Tìm bảng cover & fill ----------
def _find_result_table(doc: Document):
    cover_labels = {
        "Sample Description:", "Item/ Material code:", "Category:", "Collection:",
        "Country of Destination:", "Supplier/ Subcontractor:", "Customer:",
        "Sample Size:", "Sample Weight:", "Tested by:", "Generated by:"
    }
    for t in doc.tables:
        if len(t.columns) >= 4:
            seen = set()
            for r in t.rows:
                for c in r.cells:
                    txt = (c.text or "").strip()
                    if txt in cover_labels:
                        seen.add(txt)
            if len(seen) >= 4:
                return t
    return doc.tables[0] if doc.tables else None

def fill_bed_cover_from_excel(template_docx_path: str, excel_path_or_name: str, report_id: str) -> BytesIO:
    """
    Điền trang cover (bảng RESULT 4 cột + tiêu đề 'RESULT: -') theo Report #:
      - Supplier/ Subcontractor  <-  'QA comment' (ưu tiên)
      - RESULT: (ô '-' cạnh nhãn) <-  'Rating'
      - Test time (checkbox)     <-  'Remark' (ưu tiên tick-REGEX kiểu TRF, fallback vị trí)
      - Các trường còn lại fuzzy; tránh map sang 'QR Code' cho Item code.
      - Chỉ ghi đè khi ô value đang '-' hoặc trống; text theo font/cỡ của nhãn.
    """
    if not os.path.exists(template_docx_path):
        raise FileNotFoundError(f"Template .docx không tồn tại: {template_docx_path}")

    df = _bed_load_excel_df(excel_path_or_name)

    # cột khóa Report #
    key_col = None
    for name in ["Report #", "Report#", "Report No", "Report no", "Report", "Report_No", "Report_Number"]:
        if name in df.columns:
            key_col = name
            break
    if not key_col:
        raise KeyError("Không tìm thấy cột 'Report #' (hoặc biến thể) trong Excel.")

    row_df = df.loc[df[key_col].astype(str).str.strip() == str(report_id or "").strip()]
    if row_df.empty:
        raise ValueError(f"Không tìm thấy dòng có Report # = {report_id} trong Excel.")
    row = row_df.iloc[0]
    excel_cols = list(df.columns)

    # alias ưu tiên (đặc biệt: Supplier/Subcontractor ưu tiên QA comment)
    preferred = {
        "Result:": ["rating", "Rating", "RATING"],
        "Sample Description:": ["Item name / Description", "Item name/Description", "Description", "Item name", "Item Description"],
        "Item/ Material code:": ["Item#", "Item #", "Item code", "Item / Material code", "Item/ Material code"],
        "Category:": ["Category / Component name / Position ", "Category", "Component name", "Position"],
        "Collection:": ["Collection"],
        "Country of Destination:": ["Country of destination", "Country of Destination", "Destination"],
        "Supplier/ Subcontractor:": ["QA comment", "QA Comment", "QA comments", "Supplier / Subcontractor ", "Supplier/ Subcontractor", "Supplier", "Subcontractor"],
        "Customer:": ["Customer / Buyer", "Customer", "Buyer"],
        "Sample Size:": ["Sample Size", "Size"],
        "Sample Weight:": ["Sample Weight", "Weight"],
    }

    # cột đặc biệt
    col_rating = next((c for c in excel_cols if _bed_norm(c) == "rating"), "")
    col_remark = next((c for c in excel_cols if _bed_norm(c) == "remark"), "")

    doc = Document(template_docx_path)

    # (1) RESULT từ 'rating'
    if col_rating:
        _set_result_value(doc, _bed_val(row, col_rating))

    # (2) Bảng cover (4 cột)
    tbl = _find_result_table(doc)
    if tbl is None:
        raise RuntimeError("Không tìm thấy bảng cover/RESULT trong template.")

    for r in tbl.rows:
        cells = r.cells
        # Trái
        if len(cells) >= 2:
            l_label = (cells[0].text or "").strip()
            if l_label.endswith(":"):
                cand = _bed_pick_best_column(excel_cols, l_label, preferred_aliases=preferred.get(l_label))
                if cand:
                    cur = (cells[1].text or "").strip()
                    if _is_placeholder_dash(cur):
                        _set_cell_text_with_style(cells[1], cells[0], _bed_val(row, cand))
        # Phải
        if len(cells) >= 4:
            r_label = (cells[2].text or "").strip()
            if r_label.endswith(":"):
                cand = _bed_pick_best_column(excel_cols, r_label, preferred_aliases=preferred.get(r_label))
                if cand:
                    cur = (cells[3].text or "").strip()
                    if _is_placeholder_dash(cur):
                        _set_cell_text_with_style(cells[3], cells[2], _bed_val(row, cand))

    # (3) Tick checkbox Test time theo Remark
    if col_remark:
        raw_remark = _bed_val(row, col_remark)
        picked_label = _canon_test_label(raw_remark)  # -> '1st test' / '2nd test' / ...
        if picked_label:
            ok = _tick_testtime_by_regex(doc, picked_label)  # ƯU TIÊN: theo cơ chế TRF (regex)
            if not ok:
                _tick_testtime_checkbox_fallback(doc, picked_label)  # DỰ PHÒNG: tick ở ô bên trái

    # Xuất ra stream
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# ========================= END OF FILE =========================
