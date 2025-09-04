# testlab_dashboard.py
# -*- coding: utf-8 -*-
"""
TestLab Dashboard — read Excel from config.local_main, auto-map columns, robust late detection, rich charts.
Routes:
  - GET /testlab/dashboard            -> render template
  - GET /testlab/dashboard/data       -> JSON metrics & series (filters)
  - GET /testlab/dashboard/audit      -> mapping + coverage debug
"""

from flask import Blueprint, render_template, jsonify, request
from datetime import datetime, timedelta, date
import pandas as pd
import os, re, math

dashboard_bp = Blueprint("testlab_dashboard", __name__, template_folder="templates")

# ----------------------------------------------------------------------
# 1) READ CONFIG EXACTLY AS-IS (NO /mnt/data PREFIX)
# ----------------------------------------------------------------------
from config import local_main as DS_PATH  # <-- dùng đúng như bạn yêu cầu

# ----------------------------------------------------------------------
# 2) HELPERS
# ----------------------------------------------------------------------
def _norm(s): 
    return re.sub(r"\s+", " ", str(s if s is not None else "")).strip().lower()

def _parse_date(v):
    if v is None or str(v).strip() == "": return None
    # Excel serial
    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
        try:
            dt = pd.to_datetime(v, origin="1899-12-30", unit="D", errors="coerce")
            return None if pd.isna(dt) else dt.to_pydatetime()
        except: pass
    # pandas / datetime
    if hasattr(v, "to_pydatetime"):
        try: return v.to_pydatetime()
        except: pass
    if isinstance(v, datetime): return v
    # string
    try:
        dt = pd.to_datetime(str(v), errors="coerce", dayfirst=True)
        if pd.isna(dt):
            dt = pd.to_datetime(str(v), errors="coerce", dayfirst=False)
        return None if pd.isna(dt) else dt.to_pydatetime()
    except: 
        return None

def _month_key(dt): return dt.strftime("%Y-%m")

# ----------------------------------------------------------------------
# 3) SMART EXCEL LOADER (scan header up to 50 rows)
# ----------------------------------------------------------------------
def _read_excel_smart(path: str):
    if not os.path.exists(path):
        return pd.DataFrame(), {"path": path, "header_row": None, "columns": [], "error": "file_not_found"}
    try:
        df0 = pd.read_excel(path, sheet_name=0, header=0, engine="openpyxl")
    except Exception as e:
        return pd.DataFrame(), {"path": path, "header_row": None, "columns": [], "error": str(e)}

    bad_header = (len(df0.columns) == 0) or (
        len(df0.columns) > 0 and sum(str(c).startswith("Unnamed") for c in df0.columns) / len(df0.columns) > 0.5
    )

    if bad_header:
        raw = pd.read_excel(path, sheet_name=0, header=None, engine="openpyxl")
        best_row, best_score = 0, -1
        for r in range(min(50, len(raw))):
            row = raw.iloc[r]
            score = int(row.notna().sum()) + int(sum(isinstance(v, str) and v.strip() != "" for v in row))
            if score > best_score:
                best_score, best_row = score, r
        df = pd.read_excel(path, sheet_name=0, header=best_row, engine="openpyxl")
        df.columns = [_norm(c) for c in df.columns]
        # drop columns that are completely empty
        df = df.dropna(axis=1, how="all")
        return df, {"path": path, "header_row": best_row, "columns": list(df.columns), "error": None}
    else:
        df0.columns = [_norm(c) for c in df0.columns]
        df0 = df0.dropna(axis=1, how="all")
        return df0, {"path": path, "header_row": 0, "columns": list(df0.columns), "error": None}

# ----------------------------------------------------------------------
# 4) COLUMN ALIASES (covering your exact labels and variants/typos)
# ----------------------------------------------------------------------
ALIASES = {
    # IDs
    "trq": {"trq", "trq id", "tl-", "tl", "trq code"},
    "report_no": {"report #", "report#", "report no", "report number", "report"},
    "item_no": {"item #", "item#", "item", "item code", "code", "mã hàng"},
    # Descriptors
    "type_of": {"type of", "type_of", "type", "test group", "contruction", "construction", "transit", "material", "finishing", "outsource"},
    "country": {"country of destination", "country ò destination", "country of destinatio", "country", "destination"},
    "testing_env": {"furniture testing", "testing", "indoor/outdoor", "environment"},
    "customer": {"customer / buyer", "customer", "buyer"},
    "submitter": {"submitter in charge", "submitter", "requestor"},
    "dept": {"submitted dept.", "submitted dept", "submitted department", "department", "dept", "request dept"},
    "remark": {"remark", "lần test"},
    "qa_comment": {"qa comment", "comment", "comments"},
    "fail_reason": {"fail reason", "reason fail", "fail_reason"},
    "status": {"status", "trạng thái"},
    "rating": {"rating", "kết quả", "pass/fail", "pf", "result"},
    # Dates
    "log_in_date": {"log in date", "login date", "request date", "created date", "received date", "log in"},
    "etd": {"etd", "due", "due date", "deadline", "target date", "promise date"},
    "test_date": {"test date", "testing date"},
    "complete_date": {"complete date", "completed date", "log out date", "logout date", "finish date", "approved date"},
}

def _pick_col(cols, candidates):
    # exact then contains
    set_cols = set(cols)
    for a in candidates:
        aa = _norm(a)
        if aa in set_cols: return aa
    for a in candidates:
        aa = _norm(a)
        for c in cols:
            if aa in c: return c
    return None

def _build_mapping(df: pd.DataFrame):
    cols = list(df.columns)
    mapping = {k: _pick_col(cols, v) for k, v in ALIASES.items()}
    # coverage %
    total = len(df)
    coverage = {}
    for k, col in mapping.items():
        if col and col in df:
            coverage[k] = round(100.0 * df[col].notna().sum() / total, 1) if total else 0.0
        else:
            coverage[k] = 0.0
    return mapping, coverage

# ----------------------------------------------------------------------
# 5) STANDARDIZE & ENRICH
# ----------------------------------------------------------------------
def _std_row(r: pd.Series, m: dict):
    def g(key):
        col = m.get(key)
        if not col or col not in r: return None
        v = r[col]
        if pd.isna(v): return None
        return v
    return {
        # IDs
        "trq": str(g("trq") or "").strip(),
        "report_no": str(g("report_no") or "").strip(),
        "item_no": str(g("item_no") or "").strip(),
        # Descriptors
        "type_of": str(g("type_of") or "").strip().upper(),  # TRANSIT/CONSTRUCTION/MATERIAL/FINISHING/OUTSOURCE
        "country": str(g("country") or "").strip(),
        "testing_env": str(g("testing_env") or "").strip().upper(),  # INDOOR/OUTDOOR
        "customer": str(g("customer") or "").strip(),
        "submitter": str(g("submitter") or "").strip(),
        "dept": str(g("dept") or "").strip(),
        "remark": str(g("remark") or "").strip(),
        "qa_comment": str(g("qa_comment") or "").strip(),
        "fail_reason": str(g("fail_reason") or "").strip(),
        "status": str(g("status") or "").strip().lower(),    # done/complete/late/must/due/active...
        "rating": str(g("rating") or "").strip().lower(),    # pass/fail/data/cancel
        # Dates
        "log_in_date": _parse_date(g("log_in_date")),
        "etd": _parse_date(g("etd")),
        "test_date": _parse_date(g("test_date")),
        "complete_date": _parse_date(g("complete_date")),
    }

def _enrich_flags(rows):
    """Derive robust late/ontime flags."""
    today = date.today()
    for d in rows:
        li, etd, co = d["log_in_date"], d["etd"], d["complete_date"]
        d["is_completed"] = co is not None
        d["has_etd"] = etd is not None
        # Completed On-time/Late
        d["is_ontime_complete"] = bool(co and etd and co <= etd)
        d["is_late_complete"]   = bool(co and etd and co >  etd)
        # Overdue WIP (not yet completed, past ETD)
        d["is_overdue_wip"] = bool((not co) and etd and today > etd.date())
        # Due-today WIP
        d["is_due_today"] = bool((not co) and etd and today == etd.date())
        # Must (1 ngày trước ETD) heuristic
        d["is_must"] = bool((not co) and etd and (etd.date() - today).days == 1)
        # Status hints
        st = d["status"] or ""
        d["status_latish"] = any(k in st for k in ("late","must","due"))
        # Fail/Pass inference
        rat = (d["rating"] or "")
        d["is_fail"] = (rat == "fail") or ("fail" in (d["status"] or "")) or ("fail" in (d["fail_reason"] or "").lower())
        d["is_pass"] = (rat == "pass")
    return rows

def _load_rows_and_info():
    df, meta = _read_excel_smart(DS_PATH)
    mapping, coverage = _build_mapping(df)
    rows = [_std_row(r, mapping) for _, r in df.iterrows()]
    rows = _enrich_flags(rows)
    info = {
        "ds_meta": meta,
        "mapping": mapping,
        "coverage_pct": coverage,
        "rows_count": len(rows),
        "path_used": DS_PATH
    }
    return rows, info

# ----------------------------------------------------------------------
# 6) KPIs & SERIES
# ----------------------------------------------------------------------
def _time_range_bounds(time_range: str):
    now = datetime.now()
    if time_range == "day":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    elif time_range == "week":
        start = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    end = now.replace(hour=23, minute=59, second=59, microsecond=999999)
    return start, end

def _kpis(rows, start, end):
    total_in  = sum(1 for d in rows if d["log_in_date"] and start <= d["log_in_date"] <= end)
    total_out = sum(1 for d in rows if d["complete_date"] and start <= d["complete_date"] <= end)
    total_wip = sum(1 for d in rows if not d["complete_date"])

    # Completed with ETD
    completed_with_etd = [d for d in rows if d["complete_date"] and d["etd"]]
    ontime_completed = sum(1 for d in completed_with_etd if d["is_ontime_complete"])
    late_completed   = sum(1 for d in completed_with_etd if d["is_late_complete"])
    denom_completed  = max(1, len(completed_with_etd))
    pct_ontime = round(100.0 * ontime_completed / denom_completed, 1)
    pct_late_completed = round(100.0 * late_completed / denom_completed, 1)

    overdue_wip = sum(1 for d in rows if d["is_overdue_wip"])
    total_late = late_completed + overdue_wip  # => phản ánh cả completed late + WIP đã quá hạn

    capacity = round(total_out / max(1, (end.date() - start.date()).days + 1), 2)

    return {
        "total_in": total_in,
        "total_out": total_out,
        "capacity": capacity,
        "total_wip": total_wip,
        "total_ontime": ontime_completed,
        "total_late": total_late,
        "late_breakdown": {"completed_late": late_completed, "overdue_wip": overdue_wip},
        "pct_ontime": pct_ontime,
        "pct_late": pct_late_completed,           # %Late trên completed-with-ETD
        "pct_late_overall": round(100.0 * total_late / max(1, len(rows)), 1)  # %Late overall (tham khảo)
    }

def _series_inout_by_month(rows):
    m_in, m_out = {}, {}
    for d in rows:
        if d["log_in_date"]:
            k = _month_key(d["log_in_date"])
            m_in[k] = m_in.get(k, 0) + 1
        if d["complete_date"]:
            k = _month_key(d["complete_date"])
            m_out[k] = m_out.get(k, 0) + 1
    months = sorted(set(m_in) | set(m_out))
    return months, [m_in.get(m, 0) for m in months], [m_out.get(m, 0) for m in months]

def _series_wip_by_type(rows):
    cnt = {}
    for d in rows:
        if not d["complete_date"]:
            t = d["type_of"] or "(UNKNOWN)"
            cnt[t] = cnt.get(t, 0) + 1
    types = sorted(cnt)
    return types, [cnt[t] for t in types]

def _series_fail_ratio_by_type(rows):
    tallies = {}
    for d in rows:
        t = d["type_of"] or "(UNKNOWN)"
        if t not in tallies: tallies[t] = {"PASS": 0, "FAIL": 0}
        if d["is_pass"]: tallies[t]["PASS"] += 1
        elif d["is_fail"]: tallies[t]["FAIL"] += 1
    types, ratios = [], []
    for t in sorted(tallies):
        P, F = tallies[t]["PASS"], tallies[t]["FAIL"]
        tot = P + F
        ratios.append(round(100.0 * F / tot, 1) if tot else 0.0)
        types.append(t)
    return types, ratios

def _series_leadtime(rows):
    bucket = {}
    for d in rows:
        li = d["log_in_date"]
        if not li: continue
        key = _month_key(li)
        tgt = (d["etd"].date() - li.date()).days if d["etd"] else None
        act = (d["complete_date"].date() - li.date()).days if d["complete_date"] else None
        if key not in bucket: bucket[key] = {"t_sum":0,"t_cnt":0,"a_sum":0,"a_cnt":0}
        if tgt is not None: bucket[key]["t_sum"] += max(0, tgt); bucket[key]["t_cnt"] += 1
        if act is not None: bucket[key]["a_sum"] += max(0, act); bucket[key]["a_cnt"] += 1
    months = sorted(bucket)
    t = [round(bucket[m]["t_sum"]/bucket[m]["t_cnt"],2) if bucket[m]["t_cnt"] else 0.0 for m in months]
    a = [round(bucket[m]["a_sum"]/bucket[m]["a_cnt"],2) if bucket[m]["a_cnt"] else 0.0 for m in months]
    return months, t, a

def _series_wip_total_by_month(rows):
    """Monthly snapshot of WIP count at month end (last 12 months)."""
    # get month range
    months_set = set()
    for d in rows:
        if d["log_in_date"]: months_set.add(_month_key(d["log_in_date"]))
        if d["complete_date"]: months_set.add(_month_key(d["complete_date"]))
        if d["etd"]: months_set.add(_month_key(d["etd"]))
    months = sorted(months_set)[-12:]  # last 12 months
    # helper: month end
    def month_end(ym):
        y, m = map(int, ym.split("-"))
        if m == 12: return date(y, 12, 31)
        return (date(y, m+1, 1) - timedelta(days=1))
    totals = []
    for ym in months:
        end_day = month_end(ym)
        c = 0
        for d in rows:
            li = d["log_in_date"].date() if d["log_in_date"] else None
            co = d["complete_date"].date() if d["complete_date"] else None
            if li and li <= end_day and (co is None or co > end_day):
                c += 1
        totals.append(c)
    return months, totals

def _pf_donut(rows):
    P = sum(1 for d in rows if d["is_pass"])
    F = sum(1 for d in rows if d["is_fail"])
    O = max(0, len(rows) - P - F)
    return ["PASS","FAIL","OTHER"], [P,F,O]

# ----------------------------------------------------------------------
# 7) ROUTES
# ----------------------------------------------------------------------
@dashboard_bp.route("/dashboard")
def dashboard_page():
    rows, _ = _load_rows_and_info()
    type_set = sorted(set(d["type_of"] for d in rows if d["type_of"]))
    dept_set = sorted(set(d["dept"] for d in rows if d["dept"]))
    return render_template("dashboard.html", type_set=type_set, dept_set=dept_set)

@dashboard_bp.route("/dashboard/data")
def dashboard_data():
    # filters
    time_range = (request.args.get("time") or "month").lower()
    f_type = (request.args.get("type") or "").strip().upper()
    f_dept = (request.args.get("dept") or "").strip()
    exclude_outsource = (request.args.get("exclude_outsource") or "1") in ("1","true","True")

    rows, _ = _load_rows_and_info()

    # apply filters
    def ok(d):
        if f_type and d["type_of"] != f_type: return False
        if exclude_outsource and d["type_of"] == "OUTSOURCE": return False
        if f_dept and d["dept"] != f_dept: return False
        return True
    rows = [d for d in rows if ok(d)]

    start, end = _time_range_bounds(time_range)

    kpi = _kpis(rows, start, end)
    m_inout, v_in, v_out = _series_inout_by_month(rows)
    t_wip, v_wip = _series_wip_by_type(rows)
    t_fail, v_fail = _series_fail_ratio_by_type(rows)
    m_lt, v_tgt, v_act = _series_leadtime(rows)
    m_wip, v_wip_tot = _series_wip_total_by_month(rows)
    pf_labels, pf_vals = _pf_donut(rows)

    return jsonify({
        "kpi": kpi,
        "inout_months": m_inout, "in_by_month": v_in, "out_by_month": v_out,
        "types_wip": t_wip, "wip_by_type": v_wip,
        "types_fail": t_fail, "fail_ratio": v_fail,
        "lead_months": m_lt, "target_leadtime": v_tgt, "actual_leadtime": v_act,
        "wip_months": m_wip, "wip_total": v_wip_tot,
        "pf_labels": pf_labels, "pf_values": pf_vals,
        "filters": {"time": time_range, "type": f_type, "dept": f_dept, "exclude_outsource": exclude_outsource}
    })

@dashboard_bp.route("/dashboard/audit")
def dashboard_audit():
    rows, info = _load_rows_and_info()
    # quick missing %
    keys = ["trq","report_no","item_no","type_of","country","testing_env","customer","submitter","dept",
            "status","rating","log_in_date","etd","test_date","complete_date"]
    missing = {k: 0 for k in keys}
    for d in rows:
        for k in keys:
            v = d.get(k)
            if v is None or (isinstance(v, str) and v.strip() == ""):
                missing[k] += 1
    total = len(rows) or 1
    pct_missing = {k: round(100.0 * missing[k] / total, 1) for k in keys}

    info_out = dict(info)
    info_out["missing_pct"] = pct_missing
    return jsonify(info_out)
