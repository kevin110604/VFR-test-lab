# vfr3.py
import os, math, re, time
from datetime import datetime
from typing import Tuple, Optional

import pandas as pd
from flask import Blueprint, render_template, request, jsonify, redirect, url_for, session
from werkzeug.utils import secure_filename

# ===== Blueprint =====
vfr3_bp = Blueprint("vfr3", __name__)

# ===== Cấu hình / hằng số =====
STATIC_ROOT = "static"
DATA_ROOT   = os.path.join(STATIC_ROOT, "VFR3", "data")      # CSV mỗi khu vực
ALLOWED_IMG = {".jpg", ".jpeg", ".png", ".webp"}

# Map tên hiển thị
AREA_TITLE_MAP = {
    "vfr3/wax": "Mẫu sáp",
    "vfr3/sand-casting": "Mẫu đúc cát",
    "vfr3/ceramic-plaster": "Mẫu ceramic thạch cao",
}

# Cột chuẩn
INV_COLS = [
    "STT","Hình mẫu","Miêu tả","Vị trí","Code",
    "Số lượng Parts","Part code","Ngày tạo khuôn",
    "Ngày mượn mẫu","Người mượn","Ngày trả","Người trả",
    "Số ngày trong kho","Tuổi thọ (tháng)","Số lần đã sử dụng",
    "Số lần sử dụng tối đa","Tình trạng","Hình dạng","Ảnh mượn","Ảnh trả"
]

# ===== Helpers =====
def ensure_dirs():
    os.makedirs(DATA_ROOT, exist_ok=True)
    os.makedirs(os.path.join(STATIC_ROOT, "VFR3", "borrow"), exist_ok=True)
    os.makedirs(os.path.join(STATIC_ROOT, "VFR3", "return"), exist_ok=True)
    os.makedirs(os.path.join(STATIC_ROOT, "VFR3", "product"), exist_ok=True)

def csv_path(area_path: str) -> str:
    """Mỗi area_path 1 file, ví dụ data/VFR3/vfr3_wax.csv"""
    ensure_dirs()
    safe = area_path.replace("/", "_")
    return os.path.join(DATA_ROOT, f"{safe}.csv")

def area_name(area_path: str) -> str:
    return AREA_TITLE_MAP.get(area_path, area_path)

def load_df(area_path: str) -> pd.DataFrame:
    p = csv_path(area_path)
    if not os.path.exists(p):
        return pd.DataFrame(columns=INV_COLS)
    df = pd.read_csv(p, dtype=str, keep_default_na=False)
    for c in INV_COLS:
        if c not in df.columns:
            df[c] = ""
    # chuẩn số
    for c in ["Số lượng Parts","Số ngày trong kho","Tuổi thọ (tháng)","Số lần đã sử dụng","Số lần sử dụng tối đa"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df

def save_df(area_path: str, df: pd.DataFrame):
    if "STT" in df.columns:
        df = df.copy()
        df["STT"] = list(range(1, len(df)+1))
    p = csv_path(area_path)
    out = df.copy()
    for c in ["Số lượng Parts","Số ngày trong kho","Tuổi thọ (tháng)","Số lần đã sử dụng","Số lần sử dụng tối đa"]:
        out[c] = out[c].astype(str)
    out.to_csv(p, index=False, encoding="utf-8-sig")

def allowed_image(fn: str) -> bool:
    return os.path.splitext(fn.lower())[1] in ALLOWED_IMG

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def find_code(area_path: str, code: str) -> Tuple[Optional[dict], Optional[list], Optional[str]]:
    """
    Trả về:
      - thong_tin: dict 1 dòng (SP chính/Part) nếu tìm thấy chính xác
      - list_parts: list các part nếu code là SP chính có part
      - thong_bao: thông báo nếu không thấy
    Logic:
      - Nếu nhập đúng Part code → trả về dòng part
      - Nếu nhập Code chính:
          + nếu có part → list_parts
          + nếu không → dòng chính
    """
    df = load_df(area_path)
    if not code:
        return None, None, None

    # Tìm đúng Part code trước
    mask_part = (df["Part code"].astype(str).str.lower() == code.lower())
    if mask_part.any():
        row = df[mask_part].iloc[0].to_dict()
        return row, None, None

    # Tìm theo Code chính
    mask_code = (df["Code"].astype(str).str.lower() == code.lower())
    if not mask_code.any():
        return None, None, "Không tìm thấy mã sản phẩm."

    sub = df[mask_code]
    parts = sub[sub["Part code"].astype(str) != ""]
    if len(parts) > 0:
        return None, [r for _, r in parts.iterrows()], None
    # Không có part → trả dòng chính
    row_main = sub.iloc[0].to_dict()
    return row_main, None, None

def inc_used_safe(v_used, v_max):
    try:
        u = int(v_used) if v_used not in ("", None) else 0
    except:
        u = 0
    try:
        m = int(v_max) if v_max not in ("", None) else 0
    except:
        m = 0
    return u+1, m

# ===== MENU AREA =====
@vfr3_bp.get("/vfr3/wax")
def vfr3_wax_menu():
    return render_template("vfr3_area.html",
                           area="vfr3/wax",
                           area_name=area_name("vfr3/wax"))

@vfr3_bp.get("/vfr3/sand-casting")
def vfr3_sand_menu():
    return render_template("vfr3_area.html",
                           area="vfr3/sand-casting",
                           area_name=area_name("vfr3/sand-casting"))

@vfr3_bp.get("/vfr3/ceramic-plaster")
def vfr3_ceramic_menu():
    return render_template("vfr3_area.html",
                           area="vfr3/ceramic-plaster",
                           area_name=area_name("vfr3/ceramic-plaster"))

# ===== BORROW =====
@vfr3_bp.get("/<path:area_path>/borrow")
def vfr3_borrow(area_path):
    code = (request.args.get("code") or "").strip()
    thong_tin, list_parts, thong_bao = (None, None, None)
    if code:
        thong_tin, list_parts, thong_bao = find_code(area_path, code)

    # Chuẩn về dict list_parts
    if list_parts:
        list_parts = [r if isinstance(r, dict) else r.to_dict() for r in list_parts]

    return render_template("borrow.html",
                           area=area_path,
                           area_name=area_name(area_path),
                           thong_tin=thong_tin,
                           list_parts=list_parts,
                           thong_bao=thong_bao)

@vfr3_bp.post("/<path:area_path>/muon-xac-nhan")
def vfr3_borrow_confirm(area_path):
    try:
        code      = (request.form.get("code") or "").strip()
        part_code = (request.form.get("part_code") or "").strip()
        file_img  = request.files.get("anh_muon")

        if not code or not file_img or not file_img.filename:
            return jsonify({"success": False, "msg": "Thiếu dữ liệu cần thiết."}), 400

        if not allowed_image(file_img.filename):
            return jsonify({"success": False, "msg": "Định dạng ảnh không hợp lệ."}), 400

        # Lưu ảnh vào static/VFR3/borrow
        ensure_dirs()
        img_name = secure_filename(f"{(part_code or code)}_{int(time.time())}{os.path.splitext(file_img.filename)[1].lower()}")
        save_dir = os.path.join(STATIC_ROOT, "VFR3", "borrow")
        os.makedirs(save_dir, exist_ok=True)
        file_img.save(os.path.join(save_dir, img_name))

        # Cập nhật CSV
        df = load_df(area_path)
        if part_code:
            mask = (df["Code"].astype(str) == code) & (df["Part code"].astype(str) == part_code)
        else:
            mask = (df["Code"].astype(str) == code) & (df["Part code"].astype(str) == "")

        if not mask.any():
            return jsonify({"success": False, "msg": "Không tìm thấy sản phẩm để cập nhật."}), 404

        idx = df.index[mask][0]
        df.at[idx, "Tình trạng"]    = "Unavailable"
        df.at[idx, "Ngày mượn mẫu"] = now_str()
        df.at[idx, "Người mượn"]    = session.get("staff_id", "")
        df.at[idx, "Ảnh mượn"]      = img_name

        mieu_ta = df.at[idx, "Miêu tả"]
        vi_tri  = df.at[idx, "Vị trí"]

        save_df(area_path, df)

        # Luôn trả JSON (borrow.html sẽ popup rồi tự về /<area>)
        return jsonify({
            "success": True,
            "code": code,
            "part_code": part_code,
            "mieu_ta": mieu_ta,
            "vi_tri": vi_tri,
            "img": img_name
        }), 200

    except Exception as e:
        # Không để crash ra HTML debugger nữa
        return jsonify({"success": False, "msg": f"Lỗi xử lý: {e}"}), 500

# ===== RETURN =====
@vfr3_bp.route("/<path:area_path>/return", methods=["GET","POST"])
def vfr3_return(area_path):
    # nhận code từ query (GET) hoặc form (POST)
    code = ""
    if request.method == "POST":
        code = (request.form.get("code") or "").strip()
    else:
        code = (request.args.get("code") or "").strip()

    thong_tin, list_parts_unavailable, thong_bao = (None, None, None)
    if code:
        # lấy item / parts
        one, parts, msg = find_code(area_path, code)
        if msg:
            thong_bao = msg
        else:
            df = load_df(area_path)
            if parts:
                # lọc các part đang Unavailable
                lst = []
                for p in parts:
                    r = p if isinstance(p, dict) else p.to_dict()
                    if str(r.get("Tình trạng","")) == "Unavailable":
                        lst.append(r)
                list_parts_unavailable = lst if lst else None
                if not list_parts_unavailable:
                    thong_bao = "Tất cả parts đang khả dụng, không có part nào cần trả."
            elif one:
                # Nếu là dòng chính: chỉ cho trả khi đang Unavailable
                if str(one.get("Tình trạng","")) == "Unavailable":
                    thong_tin = one
                else:
                    thong_bao = "Sản phẩm đang khả dụng, không có gì để trả."
    return render_template("return.html",
                           area=area_path,
                           area_name=area_name(area_path),
                           thong_tin=thong_tin,
                           list_parts_unavailable=list_parts_unavailable,
                           thong_bao=thong_bao)

@vfr3_bp.post("/<path:area_path>/return/confirm")
def vfr3_return_confirm(area_path):
    code      = (request.form.get("code") or "").strip()
    part_code = (request.form.get("part_code") or "").strip()
    ke = (request.form.get("ke") or "").strip().upper()
    hang = (request.form.get("hang") or "").strip().upper()
    o = (request.form.get("o") or "").strip().upper()
    file_img  = request.files.get("anh_tra")

    if not code or not file_img or not file_img.filename:
        return jsonify({"success": False, "msg": "Thiếu dữ liệu cần thiết."})
    if not allowed_image(file_img.filename):
        return jsonify({"success": False, "msg": "Định dạng ảnh không hợp lệ."})

    ensure_dirs()
    img_name = secure_filename(f"{(part_code or code)}_{int(time.time())}{os.path.splitext(file_img.filename)[1].lower()}")
    save_dir = os.path.join(STATIC_ROOT, "VFR3", "return")
    os.makedirs(save_dir, exist_ok=True)
    file_img.save(os.path.join(save_dir, img_name))

    # Cập nhật CSV: Available, vị trí mới, ngày trả/người trả, Ảnh trả, +1 số lần đã sử dụng
    df = load_df(area_path)
    if part_code:
        mask = (df["Code"].astype(str) == code) & (df["Part code"].astype(str) == part_code)
    else:
        mask = (df["Code"].astype(str) == code) & (df["Part code"].astype(str) == "")

    if not mask.any():
        return jsonify({"success": False, "msg": "Không tìm thấy sản phẩm để cập nhật."})
    idx = df.index[mask][0]

    # tăng số lần dùng
    used_plus, max_u = inc_used_safe(df.at[idx, "Số lần đã sử dụng"], df.at[idx, "Số lần sử dụng tối đa"])
    df.at[idx, "Số lần đã sử dụng"] = used_plus

    df.at[idx, "Tình trạng"] = "Available"
    df.at[idx, "Vị trí"] = f"{ke}{hang}{o}".strip()
    df.at[idx, "Ngày trả"] = now_str()
    df.at[idx, "Người trả"] = session.get("staff_id", "")
    df.at[idx, "Ảnh trả"] = img_name

    # xoá thông tin đang mượn
    df.at[idx, "Ngày mượn mẫu"] = ""
    df.at[idx, "Người mượn"] = ""

    mieu_ta = df.at[idx, "Miêu tả"]
    save_df(area_path, df)
    return jsonify({
        "success": True,
        "code": code,
        "part_code": part_code,
        "mieu_ta": mieu_ta,
        "vi_tri": f"{ke}-{hang}-{o}",
        "img": img_name
    })

# ===== INVENTORY (rút gọn version đủ dùng cho template hiện tại) =====
def to_status_label(v: str) -> str:
    if not v: return ""
    key = v.strip().lower()
    if key == "available": return "Available"
    if key == "unavailable": return "Unavailable"
    if key in {"none","khong ton tai","không tồn tại"}: return "Không tồn tại"
    return v

def parse_ddmmyyyy(s: str):
    try:
        return datetime.strptime(s.strip(), "%d/%m/%Y")
    except Exception:
        return None

def apply_filters(df: pd.DataFrame, args: dict) -> pd.DataFrame:
    try:
        n_rules = int(args.get("n_rules","0"))
    except: n_rules = 0
    if n_rules <= 0: return df

    out = df.copy()
    for i in range(n_rules):
        field = args.get(f"r{i}_field","")
        op    = args.get(f"r{i}_op","")
        v1    = args.get(f"r{i}_v1","")

        if field == "desc" and op == "contains":
            kw = (v1 or "").strip().lower()
            if kw:
                m = out["Miêu tả"].astype(str).str.lower().str.contains(re.escape(kw), na=False)
                out = out[m]

        elif field in {"created","borrow","return"}:
            d = parse_ddmmyyyy(args.get(f"r{i}_v1_date","") or "")
            if d:
                col = {"created":"Ngày tạo khuôn","borrow":"Ngày mượn mẫu","return":"Ngày trả"}[field]
                cdt = pd.to_datetime(out[col], errors="coerce")
                if op == "before": out = out[cdt < d]
                elif op == "on":  out = out[cdt.dt.date == d.date()]
                elif op == "after": out = out[cdt > d]

        elif field in {"borrower_code","returner_code"} and op == "emp_code":
            kw = (v1 or "").strip().lower()
            if kw:
                col = "Người mượn" if field=="borrower_code" else "Người trả"
                m = out[col].astype(str).str.lower().str.contains(re.escape(kw), na=False)
                out = out[m]

        elif field in {"days","used","umax"}:
            try: num = float(v1)
            except: continue
            col = {"days":"Số ngày trong kho","used":"Số lần đã sử dụng","umax":"Số lần sử dụng tối đa"}[field]
            s = pd.to_numeric(out[col], errors="coerce")
            if op == "under": out = out[s <= num]
            elif op == "over": out = out[s >= num]

        elif field == "status" and op == "status_is":
            lab = to_status_label(v1)
            if lab:
                out = out[out["Tình trạng"].astype(str) == lab]
    return out

@vfr3_bp.get("/<path:area_path>/inventory")
def inv_page(area_path):
    df = load_df(area_path)
    q = (request.args.get("q") or "").strip()
    f_active = request.args.get("f_active") == "1"

    if f_active:
        df_view = apply_filters(df, request.args)
    else:
        df_view = df.copy()
        if q:
            kw = q.lower()
            mask = (
                df_view["Code"].astype(str).str.lower().str.contains(re.escape(kw), na=False) |
                df_view["Part code"].astype(str).str.lower().str.contains(re.escape(kw), na=False)
            )
            df_view = df_view[mask]

    try: page = int(request.args.get("page","1"))
    except: page = 1
    per_page = 30
    if q or f_active:
        total_pages = 1
        current_page = 1
        page_df = df_view
    else:
        total = len(df_view)
        total_pages = max(1, math.ceil(total / per_page))
        current_page = max(1, min(page, total_pages))
        start = (current_page-1)*per_page
        end = start + per_page
        page_df = df_view.iloc[start:end]

    cols = [c for c in INV_COLS if c in page_df.columns]
    data = page_df.fillna("").to_dict(orient="records")

    return render_template("inventory.html",
                           area=area_path,
                           area_name=area_name(area_path),
                           columns=cols,
                           data=data,
                           q=q,
                           current_page=current_page,
                           total_pages=total_pages,
                           success_message=(request.args.get("success")=="1"),
                           open_add_popup=(request.args.get("open_add")=="1"))

@vfr3_bp.post("/<path:area_path>/them-san-pham")
def inv_add(area_path):
    df = load_df(area_path)

    code = (request.form.get("code") or "").strip()
    if not code:
        return redirect(url_for("vfr3.inv_page", area_path=area_path, open_add="1"))

    dup_main = (df["Code"].astype(str) == code) & (df["Part code"].astype(str).fillna("") == "")
    if dup_main.any():
        return redirect(url_for("vfr3.inv_page", area_path=area_path, open_add="1"))

    tuoi_tho = int(request.form.get("tuoi_tho", 0) or 0)
    sld_max  = int(request.form.get("so_lan_su_dung_toi_da", 0) or 0)
    so_parts = int(request.form.get("so_luong_parts", 0) or 0)
    mieu_ta  = (request.form.get("mieu_ta") or "").strip()

    file_main = request.files.get("hinh_dang")
    if not file_main or not file_main.filename or not allowed_image(file_main.filename):
        return redirect(url_for("vfr3.inv_page", area_path=area_path, open_add="1"))

    # Lưu ảnh sản phẩm vào static/product/<area>/<code>/<file>
    rel_main = os.path.join("VFR3", "product", area_path, code, secure_filename(file_main.filename))
    abs_main = os.path.join(STATIC_ROOT, rel_main)
    os.makedirs(os.path.dirname(abs_main), exist_ok=True)
    file_main.save(abs_main)

    now = now_str()
    if so_parts > 0:
        vi_tri = ""
        hinh_mau_main = ""
    else:
        ke = (request.form.get("ke") or "").strip().upper()
        hang = (request.form.get("hang") or "").strip().upper()
        o = (request.form.get("o") or "").strip().upper()
        vi_tri = f"{ke}{hang}{o}".strip()
        hinh_mau_main = (request.form.get("hinh_mau_main") or "").strip()

    main_row = {
        "STT":"","Hình mẫu":hinh_mau_main,"Miêu tả":mieu_ta,"Vị trí":vi_tri,
        "Code":code,"Số lượng Parts":so_parts,"Part code":"",
        "Ngày tạo khuôn":now,"Ngày mượn mẫu":"","Người mượn":"","Ngày trả":"","Người trả":"",
        "Số ngày trong kho":0,"Tuổi thọ (tháng)":0 if so_parts>0 else tuoi_tho,
        "Số lần đã sử dụng":0,"Số lần sử dụng tối đa":0 if so_parts>0 else sld_max,
        "Tình trạng":"Available","Hình dạng":rel_main,"Ảnh mượn":"","Ảnh trả":""
    }
    df = pd.concat([df, pd.DataFrame([main_row])], ignore_index=True)

    if so_parts > 0:
        for i in range(1, so_parts+1):
            part_code = (request.form.get(f"part_code_{i}") or "").strip()
            part_img  = request.files.get(f"part_img_{i}")
            part_desc = (request.form.get(f"mieu_ta_part_{i}") or "").strip()
            part_shape= (request.form.get(f"hinh_mau_part_{i}") or "").strip()
            ke_p = (request.form.get(f"ke_part_{i}") or "").strip().upper()
            hang_p = (request.form.get(f"hang_part_{i}") or "").strip().upper()
            o_p = (request.form.get(f"o_part_{i}") or "").strip().upper()
            vi_tri_p = f"{ke_p}{hang_p}{o_p}".strip()

            if not part_code or not part_img or not part_img.filename or not allowed_image(part_img.filename):
                continue

            rel_part = os.path.join("VFR3", "product", area_path, code, part_code, secure_filename(part_img.filename))
            abs_part = os.path.join(STATIC_ROOT, rel_part)
            os.makedirs(os.path.dirname(abs_part), exist_ok=True)
            part_img.save(abs_part)

            part_row = {
                "STT":"","Hình mẫu":part_shape,"Miêu tả":part_desc,"Vị trí":vi_tri_p,
                "Code":code,"Số lượng Parts":0,"Part code":part_code,
                "Ngày tạo khuôn":now,"Ngày mượn mẫu":"","Người mượn":"","Ngày trả":"","Người trả":"",
                "Số ngày trong kho":0,"Tuổi thọ (tháng)":tuoi_tho,
                "Số lần đã sử dụng":0,"Số lần sử dụng tối đa":sld_max,
                "Tình trạng":"Available","Hình dạng":rel_part,"Ảnh mượn":"","Ảnh trả":""
            }
            df = pd.concat([df, pd.DataFrame([part_row])], ignore_index=True)

    save_df(area_path, df)
    return redirect(url_for("vfr3.inv_page", area_path=area_path, success="1"))

@vfr3_bp.post("/<path:area_path>/inventory/edit")
def inv_edit(area_path):
    df = load_df(area_path)
    code = (request.form.get("code") or "").strip()
    part = (request.form.get("part_code") or "").strip()

    if not code:
        return jsonify({"success": False, "msg": "Thiếu Code."})

    mask = (df["Code"].astype(str) == code) & (df["Part code"].astype(str) == part)
    if not mask.any():
        return jsonify({"success": False, "msg": "Không tìm thấy dòng cần cập nhật."})
    idx = df.index[mask][0]

    def set_if(name, key=""):
        k = key or name
        if k in request.form:
            df.at[idx, name] = (request.form.get(k) or "").strip()

    set_if("Hình mẫu")
    set_if("Miêu tả")

    ke = (request.form.get("ke") or "").strip().upper()
    hang = (request.form.get("hang") or "").strip().upper()
    o = (request.form.get("o") or "").strip().upper()
    if ke or hang or o:
        df.at[idx, "Vị trí"] = f"{ke}{hang}{o}".strip()

    if request.form.get("Ngày tạo khuôn"):
        df.at[idx, "Ngày tạo khuôn"] = request.form.get("Ngày tạo khuôn").strip()

    if request.form.get("Tuổi thọ (tháng)"):
        try: df.at[idx, "Tuổi thọ (tháng)"] = int(float(request.form.get("Tuổi thọ (tháng)")))
        except: pass
    if request.form.get("Số lần sử dụng tối đa"):
        try: df.at[idx, "Số lần sử dụng tối đa"] = int(float(request.form.get("Số lần sử dụng tối đa")))
        except: pass

    if "Tình trạng" in request.form:
        val = request.form.get("Tình trạng") or df.at[idx, "Tình trạng"]
        df.at[idx, "Tình trạng"] = val

    save_df(area_path, df)
    return jsonify({"success": True})

@vfr3_bp.post("/<path:area_path>/inventory/delete-row")
def inv_delete(area_path):
    df = load_df(area_path)
    code = (request.form.get("code") or "").strip()
    part = (request.form.get("part_code") or "").strip()

    if not code:
        return jsonify({"success": False, "msg": "Thiếu Code."})

    if not part:
        has_parts = ((df["Code"].astype(str) == code) & (df["Part code"].astype(str) != "")).any()
        if has_parts:
            return jsonify({"success": False, "msg": "Không thể xoá sản phẩm chính khi vẫn còn PART."})

    before = len(df)
    df = df[~((df["Code"].astype(str) == code) & (df["Part code"].astype(str) == part))]
    if len(df) == before:
        return jsonify({"success": False, "msg": "Không tìm thấy dòng để xoá."})

    save_df(area_path, df)
    return jsonify({"success": True})
