# -*- coding: utf-8 -*-
"""
VFR TestLab Dashboard — robust Excel mapping (row-by-row), fuzzy header matching,
YTD/year/range time windows, audit endpoints.

Blueprint is intended to be registered with url_prefix="/testlab":
    app.register_blueprint(dashboard_bp, url_prefix="/testlab")

=> Routes exposed:
    GET  /dashboard                 -> HTML (templates/dashboard.html)
    GET  /dashboard/data            -> JSON data for charts/KPIs
    GET  /dashboard/audit           -> mapping + coverage + duplicates
    GET  /dashboard/audit/headers   -> normalized headers + which key matched
    GET  /dashboard/audit/sample    -> first N normalized rows (for quick sanity-check)
"""

from flask import Blueprint, render_template, jsonify, request
from datetime import datetime, timedelta, date
import pandas as pd
import numpy as np
import os, re, unicodedata

dashboard_bp = Blueprint("testlab_dashboard", __name__, template_folder="templates")

# --------------------------- CONFIG ---------------------------
EXCEL_PATH = os.environ.get("TESTLAB_EXCEL", "ds san pham test voi qr.xlsx")
# If None -> will read first sheet by index 0 (not pandas None which returns dict)
SHEET_NAME_ENV = os.environ.get("TESTLAB_SHEET", None)

CAPACITY_PER_DAY = int(os.environ.get("TESTLAB_CAPACITY", "20"))
TARGET_LEAD_DAYS_DEFAULT = int(os.environ.get("TESTLAB_TARGET_LEAD", "7"))

OUTSOURCE_KEYWORDS = ["OUTSOURCE", "OUT SOURCE", "OUT-SOURCE"]

# hide some KPIs (comma list), e.g. TESTLAB_HIDE_KPIS="capacity,pct_late"
HIDE_KPIS = [x.strip() for x in os.environ.get("TESTLAB_HIDE_KPIS", "").split(",") if x.strip()]

# ====== (giữ) cấu hình MA (không còn dùng để tính target, nhưng giữ để backward) ======
TARGET_MA_DAYS = int(os.environ.get("TESTLAB_TARGET_MA_DAYS", "28"))
HISTORY_DAYS_LOOKBACK = int(os.environ.get("TESTLAB_HISTORY_DAYS", "120"))
WORKING_WEEKDAYS = set([0,1,2,3,4,5]) if os.environ.get("TESTLAB_WORK_SAT","1")=="1" else set([0,1,2,3,4])
FALLBACK_CAPACITY_PER_DAY = int(os.environ.get("TESTLAB_FALLBACK_CAPACITY", str(CAPACITY_PER_DAY or 20)))

# ---- Header alias dictionary (raw; will be normalized before matching)
COLUMN_MAP = {
    "report": [
        "REPORT#", "REPORT #", "REPORT NO", "REPORT NO.", "REPORT", "REPORT_NUMBER",
        "MÃ BÁO CÁO", "SỐ BÁO CÁO", "MÃ REPORT", "REPORT ID"
    ],
    "item": [
        "ITEM#", "ITEM", "ITEM CODE", "ITEM NO", "MÃ HÀNG", "MÃ", "CODE", "SKU"
    ],
    "type_of": [
        "TYPE OF", "TYPE OF TEST", "TEST TYPE", "TYPE", "LOẠI TEST", "TEST GROUP", "GROUP"
    ],
    "dept": [
        "DEPARTMENT", "BỘ PHẬN", "DEPT", "REQUESTOR DEPT", "BỘ PHẬN REQUEST"
    ],
    "result": [
        "RATING", "RESULT", "TEST RESULT", "PASS/FAIL", "KẾT QUẢ"
    ],
    "status": [
        "STATUS", "PROCESS STATUS", "STATE", "PROGRESS", "TÌNH TRẠNG", "STATUS RESULT"
    ],
    "login_date": [
        "LOG IN DATE", "LOG-IN DATE", "LOGIN DATE", "RECEIVED DATE", "IN DATE",
        "DATE IN", "NGÀY VÀO", "NGAY VAO", "NGÀY NHẬN", "NGAY NHAN"
    ],
    "etd": [
        "ETD", "DUE DATE", "EST. DATE", "EST DATE", "NGÀY HẸN", "NGAY HEN",
        "NGÀY HẸN TRẢ", "NGAY HEN TRA", "HẸN TRẢ", "HEN TRA", "HẠN HOÀN THÀNH", "HAN HOAN THANH"
    ],
    "completed": [
        "COMPLETE DATE"
    ],
    "approved": [
        "APPROVED DATE", "APPROVED", "NGÀY DUYỆT", "NGAY DUYET"
    ],
}

# ---------------------- HEADER NORMALIZATION ----------------------
def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")

def _norm_header(s: str) -> str:
    s = str(s or "")
    s = _strip_accents(s).upper()
    s = re.sub(r"[\.\,\:\;\-\_\(\)\[\]\/\\]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _find_col(df_cols, alias_list):
    norm_map = {_norm_header(c): c for c in df_cols}
    aliases = [_norm_header(a) for a in alias_list]
    for a in aliases:
        if a in norm_map:
            return norm_map[a]
    for a in aliases:
        for nc, oc in norm_map.items():
            if a and (a in nc or nc in a):
                return oc
    return None

def _find_col_with_keywords(df_cols, keywords):
    norm_map = {_norm_header(c): c for c in df_cols}
    keys = [_norm_header(k) for k in keywords]
    for k in keys:
        for nc, oc in norm_map.items():
            if k and k in nc:
                return oc
    return None

def _parse_ui_date(s):
    if not s:
        return pd.NaT
    try:
        return pd.to_datetime(s, format="%Y-%m-%d", errors="raise").normalize()
    except Exception:
        return pd.to_datetime(s, dayfirst=True, errors="coerce").normalize()

# ---------------------- DATE PARSING ----------------------
_date_patterns = [
    "%d/%m/%Y", "%d/%m/%y",
    "%d-%m-%Y", "%d-%m-%y",
    "%Y-%m-%d",
    "%d-%b-%Y", "%d-%b-%y",
    "%d-%B-%Y", "%d-%B-%y",
    "%b-%d-%Y", "%b-%d-%y",
    "%B-%d-%Y", "%B-%d-%y",
]
_dd_mon_re = re.compile(r'^(\d{1,2})[ \-\/]([A-Za-z]{3,}|[A-Za-z]+)$')
_mon_dd_re = re.compile(r'^([A-Za-z]{3,}|[A-Za-z]+)[ \-\/](\d{1,2})$')

def _try_parse_formats(s: str):
    for fmt in _date_patterns:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return pd.NaT

def _parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    s_check = str(x).strip().upper()
    if s_check in ["", "NA", "N/A", "NULL", "NONE", "-", "00/00/0000", "0"]:
        return pd.NaT
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        try:
            return pd.to_datetime(x, unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            return pd.NaT
    if isinstance(x, (pd.Timestamp, datetime, date)):
        return pd.to_datetime(x, errors="coerce")

    s = str(x).strip()
    if not s:
        return pd.NaT

    s_norm = re.sub(r"[\.]", "/", s)
    s_norm = re.sub(r"\s+", " ", s_norm)

    dt = pd.to_datetime(s_norm, dayfirst=True, errors="coerce")
    if not pd.isna(dt): return dt
    dt = pd.to_datetime(s_norm, dayfirst=False, errors="coerce")
    if not pd.isna(dt): return dt

    dt = _try_parse_formats(s_norm)
    if not pd.isna(dt): return dt

    m = _dd_mon_re.match(s_norm)
    if m:
        day = int(m.group(1)); mon_token = m.group(2)
        for pat in ("%b", "%B"):
            try:
                mon_num = pd.to_datetime(mon_token, format=pat).month
                return pd.Timestamp(year=datetime.today().year, month=mon_num, day=day)
            except Exception:
                pass
        return pd.NaT

    m = _mon_dd_re.match(s_norm)
    if m:
        mon_token = m.group(1); day = int(m.group(2))
        for pat in ("%b", "%B"):
            try:
                mon_num = pd.to_datetime(mon_token, format=pat).month
                return pd.Timestamp(year=datetime.today().year, month=mon_num, day=day)
            except Exception:
                pass
        return pd.NaT

    return pd.NaT

def _scan_and_fix_years(series, etd_series=None, comp_series=None):
    fixed = []
    last_year = None
    for i, raw in enumerate(series):
        ts = _parse_date_any(raw)
        if not pd.isna(ts):
            fixed.append(ts); last_year = ts.year; continue

        year = None
        if etd_series is not None:
            t = _parse_date_any(etd_series.iloc[i])
            if not pd.isna(t): year = t.year
        if not year and comp_series is not None:
            t = _parse_date_any(comp_series.iloc[i])
            if not pd.isna(t): year = t.year
        if not year: year = last_year or datetime.today().year

        s = str(raw).strip()
        m1 = _dd_mon_re.match(s); m2 = _mon_dd_re.match(s)
        if not (m1 or m2):
            fixed.append(pd.NaT); continue

        if m1:
            day = int(m1.group(1)); mon_token = m1.group(2)
        else:
            mon_token = m2.group(1); day = int(m2.group(2))

        mon_num = None
        for pat in ("%b", "%B"):
            try:
                mon_num = pd.to_datetime(mon_token, format=pat).month
                break
            except Exception:
                pass
        if mon_num is None:
            fixed.append(pd.NaT); continue

        try:
            ts2 = pd.Timestamp(year=year, month=mon_num, day=day)
        except Exception:
            ts2 = pd.NaT
        fixed.append(ts2)
        if not pd.isna(ts2): last_year = ts2.year
    return pd.Series(fixed, index=series.index)

def _end_of_day(dt: datetime) -> datetime:
    return dt.replace(hour=23, minute=59, second=59, microsecond=999999)

# --------------------- EXCEL LOAD ---------------------
def load_rows_from_excel():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel not found: {EXCEL_PATH}")

    sheet_to_read = 0 if SHEET_NAME_ENV is None else SHEET_NAME_ENV
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_to_read, engine="openpyxl")
    if isinstance(df, dict):
        first_key = list(df.keys())[0]
        df = df[first_key]

    df = df.dropna(how="all")

    if "Report #" in df.columns:
        df = df[~df["Report #"].isna()]

    col_use = {}
    for key, aliases in COLUMN_MAP.items():
        col_use[key] = _find_col(df.columns, aliases)

    if not col_use["login_date"]:
        col_use["login_date"] = _find_col_with_keywords(
            df.columns,
            ["LOGIN DATE", "LOG IN", "RECEIVED DATE", "DATE IN", "NGAY VAO", "NGAY NHAN", "NGAY NHAP"]
        )
    if not col_use["completed"]:
        col_use["completed"] = _find_col_with_keywords(
            df.columns,
            ["COMPLETED", "FINISH", "DATE OUT", "OUT DATE", "NGAY HOAN THANH", "NGAY TRA", "NGAY RA"]
        )
    if not col_use["etd"]:
        col_use["etd"] = _find_col_with_keywords(
            df.columns,
            ["ETD", "DUE DATE", "NGAY HEN", "NGAY HEN TRA", "HEN TRA", "HAN HOAN THANH"]
        )

    etd_col = df[col_use["etd"]] if col_use["etd"] else pd.Series([None]*len(df))
    comp_col = df[col_use["completed"]] if col_use["completed"] else pd.Series([None]*len(df))
    login_col = df[col_use["login_date"]] if col_use["login_date"] else pd.Series([None]*len(df))
    df["login_fixed"] = _scan_and_fix_years(login_col, etd_col, comp_col)

    recs = []
    for _, row in df.iterrows():
        rec = {
            "report":    str(row[col_use["report"]]).strip() if col_use["report"] else "",
            "item":      str(row[col_use["item"]]).strip() if col_use["item"] else "",
            "type_of":   str(row[col_use["type_of"]]).strip() if col_use["type_of"] else "",
            "dept":      str(row[col_use["dept"]]).strip() if col_use["dept"] else "",
            "result":    str(row[col_use["result"]]).strip().upper() if col_use["result"] else "",
            "status":    str(row[col_use["status"]]).strip().upper() if col_use.get("status") else "",
            "login_date": row["login_fixed"],
            "etd":        _parse_date_any(row[col_use["etd"]]) if col_use["etd"] else pd.NaT,
            "completed":  _parse_date_any(row[col_use["completed"]]) if col_use["completed"] else pd.NaT,
            "approved":   _parse_date_any(row[col_use["approved"]]) if col_use["approved"] else pd.NaT,
        }
        if not rec["report"] and pd.isna(rec["login_date"]) and pd.isna(rec["completed"]):
            continue
        utype = (rec["type_of"] or "").upper()
        rec["is_outsource"] = any(k in utype for k in OUTSOURCE_KEYWORDS)
        recs.append(rec)

    return recs, col_use, list(df.columns)

# --------------------- Helper ---------------------
def _daterange(d0: date, d1: date):
    cur = d0
    while cur <= d1:
        yield cur
        cur += timedelta(days=1)

# --------------------- KPI & SERIES ---------------------
def compute(all_rows, start_dt: datetime, end_dt: datetime, type_filter, dept_filter, ex_outsource: bool):
    # Filter rows theo UI
    rows = []
    tf = (type_filter or "").strip().upper()
    dfilt = (dept_filter or "").strip().upper()
    for r in all_rows:
        if tf and (r.get("type_of", "").strip().upper() != tf):
            continue
        if dfilt and (r.get("dept", "").strip().upper() != dfilt):
            continue
        if ex_outsource and r.get("is_outsource", False):
            continue
        rows.append(r)

    def in_window(ts):
        if pd.isna(ts): return False
        return (ts >= start_dt) and (ts <= end_dt)

    # IN/OUT/WIP
    total_in  = sum(1 for r in rows if (not pd.isna(r.get("login_date"))) and in_window(r["login_date"]))
    total_out = sum(1 for r in rows if (not pd.isna(r.get("completed")))  and in_window(r["completed"]))

    wip_rows = []
    for r in rows:
        ld = r["login_date"]; cd = r["completed"]
        if pd.isna(ld): continue
        if (pd.isna(cd) or cd > end_dt) and ld <= end_dt:
            wip_rows.append(r)
    total_wip = len(wip_rows)

    comp_with_etd = [r for r in rows if (not pd.isna(r["completed"]) and not pd.isna(r["etd"]))]
    total_ontime = sum(1 for r in comp_with_etd if r["completed"] <= r["etd"])
    total_late_completed = sum(1 for r in comp_with_etd if r["completed"] > r["etd"])
    wip_with_etd = [r for r in wip_rows if not pd.isna(r["etd"])]
    overdue_wip = sum(1 for r in wip_with_etd if r["etd"] < end_dt)
    total_late = total_late_completed + overdue_wip

    denom_completed = max(1, len(comp_with_etd))
    denom_overall = max(1, len(comp_with_etd) + len(wip_with_etd))

    kpi = {
        "total_in": int(total_in),
        "total_out": int(total_out),
        "capacity": int(CAPACITY_PER_DAY),
        "total_wip": int(total_wip),
        "total_ontime": int(total_ontime),
        "total_late": int(total_late),
        "pct_ontime": round(total_ontime * 100.0 / denom_completed, 1),
        "pct_late": round(total_late_completed * 100.0 / denom_completed, 1),
        "pct_late_overall": round(total_late * 100.0 / denom_overall, 1),
    }

    # Rolling capacity 30d (tham khảo)
    start30 = end_dt.date() - timedelta(days=30)
    out30 = {}
    for r in rows:
        cd = r.get("completed")
        if not pd.isna(cd):
            cdd = cd.date()
            if start30 <= cdd <= end_dt.date():
                out30[cdd] = out30.get(cdd, 0) + 1
    capacity_rolling30 = int(round(np.mean(list(out30.values())))) if out30 else 0

    # ----- monthly series -----
    months = []
    cur_m = date(start_dt.year, start_dt.month, 1)
    endm = date(end_dt.year, end_dt.month, 1)
    while cur_m <= endm:
        months.append(cur_m.strftime("%Y-%m"))
        cur_m = date(cur_m.year + (1 if cur_m.month == 12 else 0),
                     1 if cur_m.month == 12 else cur_m.month + 1, 1)
    idx = {lab: i for i, lab in enumerate(months)}
    in_by_month  = [0]*len(months)
    out_by_month = [0]*len(months)
    for r in rows:
        ld = r["login_date"]
        if not pd.isna(ld):
            key = ld.strftime("%Y-%m")
            if key in idx: in_by_month[idx[key]] += 1
        cd = r["completed"]
        if not pd.isna(cd):
            key = cd.strftime("%Y-%m")
            if key in idx: out_by_month[idx[key]] += 1

    # WIP snapshot per month end
    wip_total = []
    for lab in months:
        y, m = lab.split("-"); y = int(y); m = int(m)
        mend = datetime(y, m, 1)
        mend = (_end_of_day(datetime(y, 12, 31)) if m==12 else (_end_of_day(datetime(y, m+1, 1) - timedelta(seconds=1))))
        mend = min(mend, end_dt)
        cnt = 0
        for r in rows:
            ld = r["login_date"]; cd = r["completed"]
            if pd.isna(ld): continue
            if (pd.isna(cd) or cd > mend) and ld <= mend:
                cnt += 1
        wip_total.append(cnt)

    # WIP by type
    wip_by_type = {}
    for r in wip_rows:
        key = (r["type_of"] or "(N/A)").strip() or "(N/A)"
        wip_by_type[key] = wip_by_type.get(key, 0) + 1
    wip_by_type = dict(sorted(wip_by_type.items(), key=lambda kv: kv[1], reverse=True)[:12])
    types_wip = list(wip_by_type.keys())
    wip_by_type_vals = [wip_by_type[t] for t in types_wip]

    # PF + Fail ratio (completed in window)
    comp_rows = [r for r in rows if in_window(r["completed"])]
    pf_counts = {"PASS": 0, "FAIL": 0, "OTHER": 0}
    for r in comp_rows:
        res = (r["result"] or "").upper()
        if "PASS" in res: pf_counts["PASS"] += 1
        elif "FAIL" in res: pf_counts["FAIL"] += 1
        else: pf_counts["OTHER"] += 1

    total_by_type, fail_by_type = {}, {}
    for r in comp_rows:
        t = (r["type_of"] or "(N/A)").strip() or "(N/A)"
        total_by_type[t] = total_by_type.get(t, 0) + 1
        if "FAIL" in (r["result"] or "").upper():
            fail_by_type[t] = fail_by_type.get(t, 0) + 1
    type_items = []
    for t, tot in total_by_type.items():
        f = fail_by_type.get(t, 0)
        ratio = 0.0 if tot == 0 else round(100.0 * f / tot, 1)
        type_items.append((t, ratio))
    type_items.sort(key=lambda x: x[1], reverse=True)
    type_items = type_items[:12]
    types_fail = [t for t, _ in type_items]
    fail_ratio = [r for _, r in type_items]

    # Leadtime (actual vs target SLA)
    lead_months = months[:]
    actual_lead, target_lead = [], []
    def _target_days_for_type(tval: str) -> int:
        u = (tval or "").strip().upper()
        if "CONSTRUCTION" in u:   return 3
        if "TRANSIT" in u:        return 3
        if "ENVIRONMENTAL" in u:  return 3
        if "FINISHING" in u:      return 5
        if "MATERIAL" in u:       return 5
        return TARGET_LEAD_DAYS_DEFAULT

    for lab in lead_months:
        y, m = lab.split("-"); y = int(y); m = int(m)
        bucket = [r for r in rows
                  if (not pd.isna(r["completed"]) and not pd.isna(r["login_date"]) and
                      r["completed"].year == y and r["completed"].month == m)]
        if bucket:
            adays = [(r["completed"] - r["login_date"]).total_seconds()/86400.0 for r in bucket]
            actual_lead.append(round(float(np.mean(adays)), 1))
            tvals = [_target_days_for_type(r.get("type_of")) for r in bucket]
            target_lead.append(round(float(np.mean(tvals)), 1) if tvals else float(TARGET_LEAD_DAYS_DEFAULT))
        else:
            actual_lead.append(0.0)
            target_lead.append(float(TARGET_LEAD_DAYS_DEFAULT))

    # ====== DAILY (ETD-based target) ======
    d0 = start_dt.date()
    d1 = end_dt.date()
    days_iso = [(d.strftime("%Y-%m-%d")) for d in _daterange(d0, d1)]
    # counters
    login_counter, completed_counter, etd_counter = {}, {}, {}
    for r in rows:
        ld = r.get("login_date")
        if not pd.isna(ld) and d0 <= ld.date() <= d1:
            login_counter[ld.date()] = login_counter.get(ld.date(), 0) + 1
        cd = r.get("completed")
        if not pd.isna(cd) and d0 <= cd.date() <= d1:
            completed_counter[cd.date()] = completed_counter.get(cd.date(), 0) + 1
        etd = r.get("etd")
        if not pd.isna(etd) and d0 <= etd.date() <= d1:
            etd_counter[etd.date()] = etd_counter.get(etd.date(), 0) + 1

    in_daily, out_daily, target_out_daily, pct_achieved_daily = [], [], [], []
    for dt_ in _daterange(d0, d1):
        inn = int(login_counter.get(dt_, 0))
        out = int(completed_counter.get(dt_, 0))
        tgt = int(etd_counter.get(dt_, 0))
        in_daily.append(inn); out_daily.append(out); target_out_daily.append(tgt)
        pct_achieved_daily.append(round(100.0 * out / tgt, 1) if tgt > 0 else (100.0 if out == 0 else 0.0))

    # today / month (ETD-based)
    today_date = d1

    # Target = số mẫu có ETD = hôm nay và chưa completed
    today_target = sum(1 for r in rows
                       if not pd.isna(r.get("etd"))
                       and r["etd"].date() == today_date
                       and pd.isna(r.get("completed")))

    # Out = số mẫu completed trong hôm nay
    today_out = sum(1 for r in rows
                    if not pd.isna(r.get("completed"))
                    and r["completed"].date() == today_date)

    # Overdue = ETD < hôm nay mà chưa completed
    today_overdue = sum(1 for r in rows
                        if not pd.isna(r.get("etd"))
                        and r["etd"].date() < today_date
                        and pd.isna(r.get("completed")))

    # % progress = out / target
    if today_target > 0:
        today_progress = round(100.0 * today_out / today_target, 1)
    else:
        today_progress = 100.0 if today_out == 0 else 0.0

    # ----------- MONTH progress (MA-based) -----------
    first_day = date(end_dt.year, end_dt.month, 1)
    last_day  = end_dt.date()

    # Month target = Moving Average (3 tháng gần nhất)
    MA_N = 3  # có thể chỉnh 3 hoặc 4
    if len(out_by_month) >= MA_N:
        month_target = int(round(np.mean(out_by_month[-MA_N:])))
    elif out_by_month:
        month_target = int(round(np.mean(out_by_month)))
    else:
        month_target = CAPACITY_PER_DAY * 20  # fallback giả định

    # Month out = số lượng thực tế completed trong tháng hiện tại
    month_out = sum(1 for r in rows
                    if not pd.isna(r.get("completed"))
                    and first_day <= r["completed"].date() <= last_day)

    if month_target > 0:
        month_progress = round(100.0 * month_out / month_target, 1)
    else:
        month_progress = 100.0 if month_out == 0 else 0.0

    # Aging buckets (WIP as-of end_dt)
    def _aging_buckets(rows_wip, asof: datetime):
        labels = ["0-7d","8-14d","15-30d","31-60d",">60d"]
        bins = [0,8,15,31,61,10_000]
        if not rows_wip: return labels, [0,0,0,0,0]
        ages = []
        for r in rows_wip:
            ld = r.get("login_date")
            if pd.isna(ld): continue
            ages.append((asof - ld).days)
        ser = pd.Series(ages).astype(int)
        cat = pd.cut(ser, bins=bins, labels=labels, right=False)
        ct = cat.value_counts().reindex(labels, fill_value=0)
        return labels, ct.tolist()
    aging_labels, aging_values = _aging_buckets(wip_rows, end_dt)

    # Service level buckets (completed)
    sl_labels = ["Ontime(<=0d)", "Late 1-3d", "Late 4-7d", "Late >7d", "No ETD"]
    sl_counts = [0,0,0,0,0]
    for r in comp_rows:
        etd = r.get("etd"); cd = r.get("completed")
        if pd.isna(etd):
            sl_counts[4] += 1
        else:
            delay = (cd - etd).days
            if delay <= 0: sl_counts[0] += 1
            elif delay <= 3: sl_counts[1] += 1
            elif delay <= 7: sl_counts[2] += 1
            else: sl_counts[3] += 1

    # daily table (last 14 days) — add "in"
    start_tbl = max(d0, d1 - timedelta(days=13))
    daily_table = []
    for dt_ in _daterange(start_tbl, d1):
        idx = (dt_ - d0).days
        inn = in_daily[idx] if 0 <= idx < len(in_daily) else 0
        tgt = target_out_daily[idx] if 0 <= idx < len(target_out_daily) else 0
        out = out_daily[idx] if 0 <= idx < len(out_daily) else 0
        pct = round(100.0 * out / tgt, 1) if tgt > 0 else (100.0 if out == 0 else 0.0)
        daily_table.append({
            "date": dt_.strftime("%Y-%m-%d"),
            "in": int(inn),
            "target": int(tgt),
            "out": int(out),
            "progress_pct": pct
        })

    # lift KPI with today/month and rolling capacity
    kpi.update({
        "capacity_rolling30": int(capacity_rolling30),
        "today_target": int(today_target),
        "today_out": int(today_out),
        "today_overdue": int(today_overdue),
        "today_progress": today_progress,
        "month_target": int(month_target),
        "month_out": int(month_out),
        "month_progress": month_progress,
    })

    # Oldest unfinished samples (WIP as of end_dt)
    oldest_wip = sorted(
        [r for r in wip_rows if not pd.isna(r.get("login_date"))],
        key=lambda r: (r["login_date"], (r.get("report") or ""))
    )
    oldest_wip_table = [{
        "report_no": r.get("report") or "",
        "type": r.get("type_of") or "",
        "department": r.get("dept") or "",
        "login_date": r["login_date"].strftime("%Y-%m-%d"),
        "days_in_wip": int((end_dt.date() - r["login_date"].date()).days)
    } for r in oldest_wip]

    # Sets for filters (from all rows)
    type_set = sorted({(r.get("type_of") or "").strip() for r in all_rows if (r.get("type_of") or "").strip()})
    dept_set = sorted({(r.get("dept") or "").strip() for r in all_rows if (r.get("dept") or "").strip()})

    # Payload — giữ key cũ + bổ sung alias để HTML hiện tại dùng được
    payload = {
        "kpi": kpi,
        "hide_kpis": HIDE_KPIS,

        # MONTHLY
        "inout_months": months,
        "in_by_month": in_by_month,
        "out_by_month": out_by_month,
        # alias tách IN/OUT
        "mon_labels": months,
        "mon_in": in_by_month,
        "mon_out": out_by_month,

        # WIP snapshot
        "wip_months": months,
        "wip_total": wip_total,

        # By type
        "types_wip": types_wip,
        "wip_by_type": wip_by_type_vals,
        "types_fail": types_fail,
        "fail_ratio": fail_ratio,

        # Leadtime
        "lead_months": lead_months,
        "actual_leadtime": actual_lead,
        "target_leadtime": target_lead,

        # PF
        "pf_labels": ["PASS","FAIL","OTHER"],
        "pf_values": [pf_counts["PASS"], pf_counts["FAIL"], pf_counts["OTHER"]],

        # DAILY (alias cho HTML)
        "days": days_iso,                     # giữ theo bản cũ
        "in_daily": in_daily,
        "out_daily": out_daily,
        "target_out_daily": target_out_daily,
        "daily_labels": days_iso,
        "daily_out": out_daily,
        "daily_target": target_out_daily,
        "pct_achieved_daily": pct_achieved_daily,
        "daily_table": daily_table,
        "oldest_wip_table": oldest_wip_table,

        # Aging & SL
        "aging_labels": aging_labels,
        "aging_values": aging_values,
        "sl_labels": sl_labels,
        "sl_values": sl_counts,

        # filter sets cho lần render đầu
        "type_set": type_set,
        "dept_set": dept_set,
    }
    return payload

# ----------------------- ROUTES -----------------------
@dashboard_bp.route("/dashboard")
def view_dashboard():
    try:
        all_rows, _, _ = load_rows_from_excel()
        type_set = sorted({(r.get("type_of") or "").strip() for r in all_rows if (r.get("type_of") or "").strip()})
        dept_set = sorted({(r.get("dept") or "").strip() for r in all_rows if (r.get("dept") or "").strip()})
    except Exception:
        type_set, dept_set = [], []
    return render_template("dashboard.html", type_set=type_set, dept_set=dept_set)

def _resolve_time_window(args):
    mode = (args.get("mode") or "ytd").lower()
    year = args.get("year")
    from_s = args.get("from")
    to_s = args.get("to")

    today = datetime.today()
    if mode == "year":
        y = int(year) if (year and str(year).isdigit()) else today.year
        start = datetime(y, 1, 1)
        end = _end_of_day(min(datetime(y, 12, 31), today))
        return mode, start, end
    elif mode == "range":
        start = _parse_ui_date(from_s) if from_s else None
        end   = _parse_ui_date(to_s) if to_s else None
        if not start or pd.isna(start): start = datetime(today.year, 1, 1)
        if not end   or pd.isna(end):   end   = today
        return mode, start, _end_of_day(end)
    else:  # ytd
        start = datetime(today.year, 1, 1)
        end = _end_of_day(today)
        return "ytd", start, end

@dashboard_bp.route("/dashboard/data")
def data_dashboard():
    try:
        all_rows, col_map, headers = load_rows_from_excel()
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    mode, start_dt, end_dt = _resolve_time_window(request.args)
    type_filter = request.args.get("type") or ""
    dept_filter = request.args.get("dept") or ""
    exclude_outsource = (request.args.get("exclude_outsource") == "1")

    js = compute(all_rows, start_dt, end_dt, type_filter, dept_filter, exclude_outsource)
    js["time"] = {"mode": mode, "start": start_dt.strftime("%Y-%m-%d"), "end": end_dt.strftime("%Y-%m-%d")}
    js["config"] = {
        "capacity_per_day": CAPACITY_PER_DAY,
        "target_lead_default": TARGET_LEAD_DAYS_DEFAULT,
        "target_ma_days": TARGET_MA_DAYS,
        "history_days": HISTORY_DAYS_LOOKBACK,
        "working_weekdays": sorted(list(WORKING_WEEKDAYS)),
    }
    js["audit"] = {
        "excel_path": EXCEL_PATH,
        "headers": headers,
        "col_map": {k: (v if v is None else str(v)) for k, v in col_map.items()}
    }
    return jsonify(js)

@dashboard_bp.route("/dashboard/audit")
def audit_dashboard():
    try:
        rows, col_map, headers = load_rows_from_excel()
        rep_counts = {}
        for r in rows:
            key = (r.get("report") or "").strip()
            if not key: continue
            rep_counts[key] = rep_counts.get(key, 0) + 1
        dups = sorted([k for k, v in rep_counts.items() if v > 1])
        return jsonify({
            "excel_path": EXCEL_PATH,
            "sheet": SHEET_NAME_ENV if SHEET_NAME_ENV is not None else 0,
            "headers": headers,
            "column_mapping": col_map,
            "total_rows_after_norm": len(rows),
            "duplicate_report_keys": dups
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@dashboard_bp.route("/dashboard/audit/headers")
def audit_headers():
    try:
        sheet_to_read = 0 if SHEET_NAME_ENV is None else SHEET_NAME_ENV
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_to_read, engine="openpyxl")
        if isinstance(df, dict):
            first_key = list(df.keys())[0]
            df = df[first_key]
        headers = list(df.columns)
        return jsonify({
            "headers_raw": headers,
            "headers_normalized": [{ "raw": h, "norm": _norm_header(h) } for h in headers]
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@dashboard_bp.route("/dashboard/audit/sample")
def audit_sample():
    try:
        limit = int(request.args.get("limit", "5"))
        rows, _, _ = load_rows_from_excel()
        return jsonify({"sample": rows[:max(0, limit)]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500