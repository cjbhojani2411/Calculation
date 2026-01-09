import os
import re
import difflib
import pandas as pd
from datetime import date, timedelta

# =========================
# CONFIG (UPDATE PATHS)
# =========================
TOPTRACKER_FILE  = "/Users/pardypanda/Documents/PPS/source/toptracker_2025_12_31_01_29.csv"
LEAVE_FILE       = "/Users/pardypanda/Documents/PPS/source/Leave View.xls"
RESOURCE_FILE    = "/Users/pardypanda/Documents/PPS/source/Resource_availability.csv"

# ✅ NEW: biometric attendance file (monthinout)
ATTENDANCE_FILE  = "/Users/pardypanda/Documents/PPS/source/monthinout02012026154043.xls"

OUTPUT_FINAL = "/Users/pardypanda/Documents/PPS/Output/monthly_payroll_summary.csv"

OUTPUT_WORKING_CALENDAR = "/Users/pardypanda/Documents/PPS/Output/_debug_working_calendar.csv"
OUTPUT_TOPTRACKER_MONTH = "/Users/pardypanda/Documents/PPS/Output/_debug_toptracker_monthly.csv"
OUTPUT_LEAVE_MONTH      = "/Users/pardypanda/Documents/PPS/Output/_debug_leave_monthly.csv"

# ✅ NEW: biometric debug outputs
OUTPUT_BIOMETRIC_DAILY  = "/Users/pardypanda/Documents/PPS/Output/_debug_biometric_daily.csv"
OUTPUT_BIOMETRIC_MONTH  = "/Users/pardypanda/Documents/PPS/Output/_debug_biometric_monthly.csv"

# NEW: name mismatch + alias files (TopTracker has no PPS code, so we map names -> PPS)
ALIASES_FILE           = "/Users/pardypanda/Documents/PPS/Output/_name_aliases.csv"
OUTPUT_NAME_MISMATCH   = "/Users/pardypanda/Documents/PPS/Output/_debug_name_mismatch.csv"
FUZZY_THRESHOLD        = 0.86  # safe default (0.0 - 1.0)

HOURS_PER_DAY = 8

# =========================
# EXCLUSIONS (case-insensitive)
# =========================
EXCLUDE_NAMES = {
    "tanay shah",
    "devang shah",
    "accounts pardy panda",
}

# =========================
# Helpers
# =========================
MONTHS = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
}

WFH_KEYWORDS = [
    "work from home", "work-from-home", "wfh", "remote work", "remote", "home office"
]

STOPWORDS = {"mohammad", "md", "mohd", "mr", "ms", "mrs", "shri", "smt"}

def norm_name(name: str) -> str:
    if pd.isna(name):
        return ""
    return " ".join(str(name).strip().lower().split())

def is_excluded_name(name: str) -> bool:
    return norm_name(name) in EXCLUDE_NAMES

def norm_tokens(name: str):
    if pd.isna(name):
        return []
    tokens = [t for t in re.split(r"\s+", str(name).strip().lower()) if t]
    return [t for t in tokens if t not in STOPWORDS]

def canonical_name(name: str) -> str:
    return " ".join(norm_tokens(name))

def best_fuzzy_match(query: str, candidates: dict) -> tuple[str, float]:
    q = canonical_name(query)
    if not q:
        return "", 0.0
    best_key, best_score = "", 0.0
    for cand in candidates.keys():
        score = difflib.SequenceMatcher(None, q, cand).ratio()
        if score > best_score:
            best_score = score
            best_key = cand
    return best_key, best_score

def infer_year_month_from_toptracker(tt: pd.DataFrame) -> tuple[int, int]:
    tt["start_time"] = pd.to_datetime(tt["start_time"], errors="coerce")
    tt = tt.dropna(subset=["start_time"])
    if tt.empty:
        raise ValueError("TopTracker has no valid start_time rows to infer payroll month.")
    dt0 = tt["start_time"].min()
    return int(dt0.year), int(dt0.month)

def parse_resource_day_column(col: str):
    m = re.match(r"^\s*(\d{1,2})\s+([A-Za-z]{3})\s*\(", str(col))
    if not m:
        return None
    d = int(m.group(1))
    mon = MONTHS.get(m.group(2))
    if not mon:
        return None
    return d, mon

def is_weekend_or_holiday(cell) -> bool:
    if pd.isna(cell):
        return False
    s = str(cell).strip().lower()
    return s == "weekend" or s.startswith("holiday")

def safe_num(x, default=0.0):
    try:
        if pd.isna(x):
            return default
        return float(x)
    except Exception:
        return default

def normalize_employee_id(emp_id: str) -> tuple[str, str]:
    parts = str(emp_id).strip().split()
    if len(parts) < 2:
        return str(emp_id).strip(), str(emp_id).strip()
    return " ".join(parts[:-1]).strip(), parts[-1].strip()

def is_wfh_leave_type(leave_type: str) -> bool:
    lt = str(leave_type).strip().lower()
    return any(k in lt for k in WFH_KEYWORDS)

def business_days_in_range(start_d: date, end_d: date) -> list[date]:
    out, cur = [], start_d
    while cur <= end_d:
        if cur.weekday() < 5:
            out.append(cur)
        cur += timedelta(days=1)
    return out

# =========================
# ✅ NEW: Biometric helpers (monthinout)
# =========================
def empcode_to_pps(empcode: str) -> str:
    """0072 -> PPS072"""
    s = str(empcode).strip()
    s = re.sub(r"[^\d]", "", s)
    if not s:
        return ""
    return f"PPS{int(s):03d}"

def hhmm_to_minutes(x) -> int:
    """Convert Work+OT cell into minutes.

    Supports:
      - '07:20' (HH:MM)
      - Excel time fractions (e.g., 0.305555 -> 7:20)
      - Decimal hours as string/number (e.g., '7.33' -> 7.33 hours)
    """
    if pd.isna(x):
        return 0

    # Excel time fractions often arrive as floats (fraction of a day)
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        val = float(x)
        if val <= 0:
            return 0
        # Heuristic: if <= 1.5 treat as fraction of a day, else treat as hours
        if val <= 1.5:
            return int(round(val * 24 * 60))
        return int(round(val * 60))

    s = str(x).strip()
    if not s or s in {"--:--", "0", "0:0", "0:00"}:
        return 0

    # HH:MM
    m = re.match(r"^(\d{1,3}):(\d{1,2})$", s)
    if m:
        hh = int(m.group(1))
        mm = int(m.group(2))
        return hh * 60 + mm

    # Decimal hours (e.g., 7.5)
    if re.match(r"^\d+(\.\d+)?$", s):
        return int(round(float(s) * 60))

    return 0

def parse_monthinout(att_path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Parses monthinout*.xls which has repeated blocks:
      Empcode | 0072 | ... | Name | ...
      Date | Shift | IN | Out | Work+OT | OT | ...
      (daily rows...)
      Total Work+OT Hrs | ... | 165:40  (or similar)
      (next Empcode block...)

    Returns:
      - daily_df: per-day rows with work_ot_minutes
      - totals_df: per-employee monthly totals from "Total Work+OT Hrs" row when present
                  (preferred over summing daily rows because some sheets omit/differ in daily structure)
    """
    if not os.path.exists(att_path):
        empty_daily = pd.DataFrame(columns=["employee_code", "employee_name_raw", "date", "in_time", "out_time", "work_ot", "work_ot_minutes"])
        empty_totals = pd.DataFrame(columns=["employee_code", "biometric_total_minutes"])
        return empty_daily, empty_totals

    try:
        raw = pd.read_excel(att_path, header=None, engine="xlrd")
    except Exception:
        # Fallback: some exports open with openpyxl even if extension is .xls
        raw = pd.read_excel(att_path, header=None, engine="openpyxl")

    rows: list[dict] = []
    totals_minutes: dict[str, int] = {}

    i = 0
    n = len(raw)

    def row_contains_total_label(row_vals: list[str]) -> bool:
        joined = " ".join([v for v in row_vals if v]).lower()
        return "total work+ot" in joined or "total work + ot" in joined or "total work+ot hrs" in joined

    while i < n:
        # Look for 'Empcode' label anywhere in the row (some exports shift columns)
        emp_label_col = None
        for col in range(raw.shape[1]):
            cell = "" if pd.isna(raw.iat[i, col]) else str(raw.iat[i, col]).strip()
            if cell.lower().replace(" ", "") in ("empcode", "empcode:", "empcode-") or cell.lower().replace(" ", "") == "empcode.":
                emp_label_col = col
                break

        if emp_label_col is not None:
            # Empcode value is usually the next non-empty cell to the right
            empcode = ""
            for col in range(emp_label_col + 1, raw.shape[1]):
                cell = "" if pd.isna(raw.iat[i, col]) else str(raw.iat[i, col]).strip()
                if cell != "":
                    empcode = cell
                    break

            employee_code = empcode_to_pps(empcode)

            # Try to read employee name from the same Empcode row (some exports have inconsistent Empcode)
            emp_name = ""
            for col in range(raw.shape[1]):
                cell = "" if pd.isna(raw.iat[i, col]) else str(raw.iat[i, col]).strip()
                if cell.lower() in {"name", "employee name", "emp name"}:
                    # take next non-empty cell to the right
                    for k in range(col + 1, raw.shape[1]):
                        c2 = "" if pd.isna(raw.iat[i, k]) else str(raw.iat[i, k]).strip()
                        if c2 != "":
                            emp_name = c2
                            break
                    break

            # header row is next line
            header_row = i + 1
            if header_row >= n:
                break

            headers = raw.iloc[header_row].tolist()
            headers = [("" if pd.isna(h) else str(h).strip()) for h in headers]

            def find_idx(name: str):
                try:
                    return headers.index(name)
                except ValueError:
                    return None

            idx_date = find_idx("Date")
            idx_in   = find_idx("IN")
            idx_out  = find_idx("Out")
            idx_work = find_idx("Work+OT")

            j = header_row + 1
            while j < n:
                # next block starts
                c0j = "" if pd.isna(raw.iat[j, 0]) else str(raw.iat[j, 0]).strip()
                if c0j.lower() == "empcode":
                    break

                # detect "Total Work+OT Hrs" row anywhere in the block
                row_vals = [("" if pd.isna(v) else str(v).strip()) for v in raw.iloc[j].tolist()]
                if row_contains_total_label(row_vals):
                    # Prefer the value from the Work+OT column if present, else take the last HH:MM-like token in the row
                    candidate = ""
                    if idx_work is not None and idx_work < len(row_vals):
                        candidate = row_vals[idx_work]
                    if not candidate:
                        # pick last HH:MM
                        for v in reversed(row_vals):
                            if re.match(r"^\d{1,3}:\d{2}$", v):
                                candidate = v
                                break
                    total_min = hhmm_to_minutes(candidate)
                    # If multiple totals appear, keep the max (safest)
                    prev = totals_minutes.get(employee_code, 0)
                    totals_minutes[employee_code] = max(prev, total_min)
                    j += 1
                    continue

                # parse daily row (needs a parsable date in Date column)
                if idx_date is None:
                    j += 1
                    continue

                cell_date = raw.iat[j, idx_date]
                try:
                    d = pd.to_datetime(cell_date, errors="raise", dayfirst=True)

                except Exception:
                    j += 1
                    continue

                in_time  = raw.iat[j, idx_in]  if idx_in  is not None else None
                out_time = raw.iat[j, idx_out] if idx_out is not None else None
                work_ot  = raw.iat[j, idx_work] if idx_work is not None else None

                work_ot_str = "" if pd.isna(work_ot) else str(work_ot).strip()
                work_ot_min = hhmm_to_minutes(work_ot_str)

                rows.append({
                    "employee_code": employee_code,
                    "employee_name_raw": emp_name,
                    "date": d.date(),
                    "in_time": "" if pd.isna(in_time) else str(in_time).strip(),
                    "out_time": "" if pd.isna(out_time) else str(out_time).strip(),
                    "work_ot": work_ot_str,
                    "work_ot_minutes": work_ot_min,
                })
                j += 1

            i = j
        else:
            i += 1

    daily_df = pd.DataFrame(rows)
    totals_df = pd.DataFrame(
        [{"employee_code": k, "biometric_total_minutes": v} for k, v in totals_minutes.items()]
    )
    return daily_df, totals_df

# =========================
# STEP 1: LOAD RESOURCE AVAILABILITY (WORKING DAYS BASE)
# =========================
ra = pd.read_csv(RESOURCE_FILE)

required_ra = {"Employee Id", "Employee Name"}
missing = required_ra - set(ra.columns)
if missing:
    raise ValueError(f"Resource_availability.csv missing columns: {missing}")

ra["employee_name_raw"] = ra["Employee Name"].astype(str).str.strip()

# ✅ EXCLUDE from Resource
ra = ra[~ra["employee_name_raw"].apply(is_excluded_name)].copy()

ra["employee_key"] = ra["employee_name_raw"].apply(norm_name)
ra["employee_code"] = ra["Employee Id"].astype(str).str.strip()

day_cols = [c for c in ra.columns if parse_resource_day_column(c)]
if not day_cols:
    raise ValueError("No day columns found in Resource_availability.csv (expected like '01 Dec (Mon)').")

resource_candidates = {}
for _, r in ra.iterrows():
    cname = canonical_name(r["employee_name_raw"])
    if cname:
        resource_candidates[cname] = str(r["employee_code"]).strip()

key_to_code = (
    ra[["employee_key", "employee_code"]]
    .dropna()
    .drop_duplicates(subset=["employee_key"])
    .set_index("employee_key")["employee_code"]
    .to_dict()
)

alias_map = {}
if os.path.exists(ALIASES_FILE):
    alias_df = pd.read_csv(ALIASES_FILE)
    if {"toptracker_name", "employee_code"}.issubset(alias_df.columns):
        alias_df["toptracker_name"] = alias_df["toptracker_name"].astype(str)
        alias_df["employee_code"] = alias_df["employee_code"].astype(str)
        alias_map = dict(zip(alias_df["toptracker_name"].apply(norm_name), alias_df["employee_code"].str.strip()))

# =========================
# STEP 2: LOAD TOPTRACKER (map to employee_code reliably)
# =========================
tt = pd.read_csv(TOPTRACKER_FILE)

required_tt = {"workers", "start_time", "duration_seconds"}
missing = required_tt - set(tt.columns)
if missing:
    raise ValueError(f"TopTracker CSV missing columns: {missing}")

tt["employee_name_raw"] = tt["workers"].astype(str).str.strip()
tt = tt[~tt["employee_name_raw"].apply(is_excluded_name)].copy()

year, month = infer_year_month_from_toptracker(tt)

tt["start_time"] = pd.to_datetime(tt["start_time"], errors="coerce")
tt = tt.dropna(subset=["start_time"]).copy()
tt = tt[(tt["start_time"].dt.year == year) & (tt["start_time"].dt.month == month)].copy()

tt["employee_key"] = tt["employee_name_raw"].apply(norm_name)

tt["duration_seconds"] = pd.to_numeric(tt["duration_seconds"], errors="coerce").fillna(0)

if "screenshot_count" in tt.columns:
    tt["screenshot_count"] = pd.to_numeric(tt["screenshot_count"], errors="coerce").fillna(0)
else:
    tt["screenshot_count"] = 0

tt["employee_code"] = tt["employee_key"].map(key_to_code)
tt.loc[tt["employee_code"].isna(), "employee_code"] = tt.loc[
    tt["employee_code"].isna(), "employee_key"
].map(alias_map)

unmatched = tt[tt["employee_code"].isna()].copy()
suggestions = []

for idx, row in unmatched.iterrows():
    original = row["employee_name_raw"]
    best_cand, score = best_fuzzy_match(original, resource_candidates)
    if score >= FUZZY_THRESHOLD and best_cand in resource_candidates:
        tt.at[idx, "employee_code"] = resource_candidates[best_cand]
    else:
        suggestions.append({
            "toptracker_name": original,
            "normalized": canonical_name(original),
            "best_match_resource": best_cand,
            "best_match_code": resource_candidates.get(best_cand, ""),
            "score": round(score, 3),
        })

if suggestions:
    pd.DataFrame(suggestions).to_csv(OUTPUT_NAME_MISMATCH, index=False)

tt = tt.dropna(subset=["employee_code"]).copy()
tt["employee_code"] = tt["employee_code"].astype(str).str.strip()

toptracker_month = (
    tt.groupby(["employee_code"], as_index=False)
      .agg(
          screenshot_count=("screenshot_count", "sum"),
          total_tracked_seconds=("duration_seconds", "sum"),
      )
)

toptracker_month["top_tracker_hours"] = (toptracker_month["screenshot_count"] / 4.0).round(2)
toptracker_month["screenshot_count"] = toptracker_month["screenshot_count"].round(0).astype(int)
toptracker_month["total_tracked_seconds"] = toptracker_month["total_tracked_seconds"].round(0).astype(int)

# =========================
# STEP 3: WORKING DAYS COUNT FROM RESOURCE
# =========================
cal_rows = []
for _, row in ra.iterrows():
    emp_code = row["employee_code"]
    emp_name_raw = row["employee_name_raw"]

    for c in day_cols:
        parsed = parse_resource_day_column(c)
        if parsed is None:
            continue

        d, m = parsed
        if m != month:
            continue
        if is_weekend_or_holiday(row.get(c)):
            continue

        cal_rows.append({
            "employee_code": emp_code,
            "employee_name": emp_name_raw,
            "date": date(year, month, d),
        })

calendar_df = pd.DataFrame(cal_rows).drop_duplicates(subset=["employee_code", "date"])

working_days = (
    calendar_df.groupby(["employee_code", "employee_name"], as_index=False)
               .agg(total_working_days=("date", "nunique"))
)

# =========================
# STEP 4: LOAD LEAVES
# =========================
lv = pd.read_excel(LEAVE_FILE, engine="xlrd")

required_lv = {"Employee ID", "From", "To", "Leave type", "Days/Hours Taken", "Approval Status"}
missing = required_lv - set(lv.columns)
if missing:
    raise ValueError(f"Leave View missing columns: {missing}")

lv = lv[lv["Approval Status"].astype(str).str.lower() == "approved"].copy()

lv["employee_name_raw"], lv["employee_code"] = zip(*lv["Employee ID"].apply(normalize_employee_id))
lv = lv[~lv["employee_name_raw"].apply(is_excluded_name)].copy()
lv["employee_code"] = lv["employee_code"].astype(str).str.strip()

leave_rows = []
for _, r in lv.iterrows():
    start = pd.to_datetime(r["From"], errors="coerce")
    end = pd.to_datetime(r["To"], errors="coerce")
    if pd.isna(start) or pd.isna(end):
        continue

    start_d = start.date()
    end_d = end.date()
    if end_d < start_d:
        continue

    leave_type = str(r["Leave type"]).strip()
    lt = leave_type.lower()

    is_unpaid = ("without" in lt) or ("lop" in lt) or ("unpaid" in lt)
    is_wfh = is_wfh_leave_type(leave_type)

    units_total = safe_num(r.get("Days/Hours Taken"), 0.0)
    if units_total <= 0:
        continue

    if (end_d.year, end_d.month) < (year, month) or (start_d.year, start_d.month) > (year, month):
        continue

    if start_d.year == year and start_d.month == month and end_d.year == year and end_d.month == month:
        leave_rows.append({
            "employee_code": r["employee_code"],
            "units": float(units_total),
            "is_unpaid": bool(is_unpaid),
            "is_wfh": bool(is_wfh),
        })
        continue

    biz_days_all = business_days_in_range(start_d, end_d)
    if not biz_days_all:
        continue

    biz_days_in_month = [d for d in biz_days_all if d.year == year and d.month == month]
    if not biz_days_in_month:
        continue

    proportion = len(biz_days_in_month) / len(biz_days_all)
    month_units = float(units_total) * proportion

    leave_rows.append({
        "employee_code": r["employee_code"],
        "units": float(round(month_units, 4)),
        "is_unpaid": bool(is_unpaid),
        "is_wfh": bool(is_wfh),
    })

leave_df = pd.DataFrame(leave_rows)

if leave_df.empty:
    leave_month = pd.DataFrame(columns=["employee_code", "leave_taken", "leave_without_paid", "work_from_home"])
else:
    non_wfh = leave_df[leave_df["is_wfh"] == False].copy()
    wfh = leave_df[leave_df["is_wfh"] == True].copy()

    leave_taken = non_wfh.groupby(["employee_code"], as_index=False).agg(leave_taken=("units", "sum"))
    leave_unpaid = non_wfh[non_wfh["is_unpaid"] == True].groupby(["employee_code"], as_index=False).agg(leave_without_paid=("units", "sum"))
    wfh_days = wfh.groupby(["employee_code"], as_index=False).agg(work_from_home=("units", "sum"))

    leave_month = leave_taken.merge(leave_unpaid, on=["employee_code"], how="left") \
                             .merge(wfh_days, on=["employee_code"], how="left")

    leave_month["leave_taken"] = leave_month["leave_taken"].fillna(0.0).round(2)
    leave_month["leave_without_paid"] = leave_month["leave_without_paid"].fillna(0.0).round(2)
    leave_month["work_from_home"] = leave_month["work_from_home"].fillna(0.0).round(2)

# =========================
# ✅ NEW STEP 4.5: LOAD BIOMETRIC ATTENDANCE (monthinout)
#   We prefer the sheet's "Total Work+OT Hrs" when available, otherwise we fall back to summing daily Work+OT.
# =========================
biometric_daily, biometric_totals = parse_monthinout(ATTENDANCE_FILE)

# keep only employees we consider + the payroll month (for debug daily export)
if not biometric_daily.empty:
    biometric_daily = biometric_daily[biometric_daily["employee_code"].isin(set(ra["employee_code"]))].copy()
    biometric_daily["date_dt"] = pd.to_datetime(biometric_daily["date"], errors="coerce")
    biometric_daily = biometric_daily[
        (biometric_daily["date_dt"].dt.year == year) &
        (biometric_daily["date_dt"].dt.month == month)
    ].copy()
    biometric_daily.drop(columns=["date_dt"], inplace=True)

# daily -> monthly minutes (fallback)
biometric_month_from_daily = (
    biometric_daily.groupby("employee_code", as_index=False)
    .agg(biometric_minutes_daily=("work_ot_minutes", "sum"))
) if not biometric_daily.empty else pd.DataFrame(columns=["employee_code", "biometric_minutes_daily"])

# totals row -> monthly minutes (preferred)
if biometric_totals is None or biometric_totals.empty:
    biometric_month_from_total = pd.DataFrame(columns=["employee_code", "biometric_minutes_total"])
else:
    biometric_month_from_total = biometric_totals.copy()
    biometric_month_from_total = biometric_month_from_total[biometric_month_from_total["employee_code"].isin(set(ra["employee_code"]))].copy()
    biometric_month_from_total.rename(columns={"biometric_total_minutes": "biometric_minutes_total"}, inplace=True)

# combine: use Total row when it exists and is > 0, else fallback to daily sum
biometric_month = biometric_month_from_daily.merge(
    biometric_month_from_total, on="employee_code", how="outer"
)

biometric_month["biometric_minutes_daily"] = biometric_month["biometric_minutes_daily"].fillna(0).astype(int)
biometric_month["biometric_minutes_total"] = biometric_month["biometric_minutes_total"].fillna(0).astype(int)

biometric_month["biometric_minutes"] = biometric_month.apply(
    lambda r: r["biometric_minutes_total"] if r["biometric_minutes_total"] > 0 else r["biometric_minutes_daily"],
    axis=1,
)

biometric_month["biometric_total_hours"] = (biometric_month["biometric_minutes"] / 60.0).round(2)

# conve
# --- right before selecting biometric_month columns ---
if "employee_name_raw" not in biometric_month.columns:
    biometric_month["employee_name_raw"] = ""

biometric_month = biometric_month[["employee_code", "employee_name_raw", "biometric_total_hours"]]


# If Empcode blocks are inconsistent/missing, prefer matching by employee name (same strategy as TopTracker)
if "employee_name_raw" in biometric_month.columns:
    biometric_month["employee_key"] = biometric_month["employee_name_raw"].apply(norm_name)
    # try direct match to resource availability
    biometric_month.loc[biometric_month["employee_code"].isna() | (biometric_month["employee_code"] == ""), "employee_code"] = biometric_month.loc[
        biometric_month["employee_code"].isna() | (biometric_month["employee_code"] == ""), "employee_key"
    ].map(key_to_code)
    # try alias map (reuse _name_aliases.csv; add biometric names in toptracker_name column)
    biometric_month.loc[biometric_month["employee_code"].isna(), "employee_code"] = biometric_month.loc[
        biometric_month["employee_code"].isna(), "employee_key"
    ].map(alias_map)

    # fuzzy match remaining
    bio_unmatched = biometric_month[biometric_month["employee_code"].isna()].copy()
    bio_suggestions = []
    for idx, row in bio_unmatched.iterrows():
        original = row.get("employee_name_raw", "")
        best_key, score = best_fuzzy_match(str(original), resource_candidates)
        suggested_code = resource_candidates.get(best_key, "") if score >= FUZZY_THRESHOLD else ""
        bio_suggestions.append((idx, best_key, score, suggested_code))
    for idx, best_key, score, suggested_code in bio_suggestions:
        if suggested_code:
            biometric_month.at[idx, "employee_code"] = suggested_code

    # export biometric name mismatches separately (safe debugging)
    bio_unmatched2 = biometric_month[biometric_month["employee_code"].isna()].copy()
    if not bio_unmatched2.empty:
        bio_unmatched2.to_csv(OUTPUT_BIOMETRIC_MONTH.replace(".csv","_name_mismatch.csv"), index=False)

# keep only resolved codes and month totals
biometric_month = biometric_month.dropna(subset=["employee_code"])
biometric_month = biometric_month[["employee_code", "biometric_total_hours"]]

# =========================
# STEP 5: FINAL MERGE + CALCS
# =========================
final = working_days.merge(toptracker_month, on=["employee_code"], how="left") \
                    .merge(leave_month, on=["employee_code"], how="left") \
                    .merge(biometric_month, on=["employee_code"], how="left")   # ✅ NEW merge

final["top_tracker_hours"] = final.get("top_tracker_hours", 0.0).fillna(0.0)
final["screenshot_count"] = final.get("screenshot_count", 0).fillna(0).astype(int)
final["total_tracked_seconds"] = final.get("total_tracked_seconds", 0).fillna(0).astype(int)

final["leave_taken"] = final.get("leave_taken", 0.0).fillna(0.0)
final["leave_without_paid"] = final.get("leave_without_paid", 0.0).fillna(0.0)
final["work_from_home"] = final.get("work_from_home", 0.0).fillna(0.0)

# ✅ NEW: biometric total hours (decimal)
#final["biometric_total_hours"] = final.get("biometric_total_hours", 0.0).fillna(0.0)
# ✅ NEW: biometric total hours (raw from biometric sheet)
final["biometric_total_hours"] = final.get("biometric_total_hours", 0.0).fillna(0.0)

# ✅ ADD: WFH hours (WFH days × 8) into biometric_total_hours
final["wfh_hours"] = (final["work_from_home"].fillna(0.0) * HOURS_PER_DAY).round(2)

# (Optional but recommended) keep original biometric value for debugging
final["biometric_total_hours_raw"] = final["biometric_total_hours"]

# Final biometric total hours = biometric + WFH hours
final["biometric_total_hours"] = (final["biometric_total_hours_raw"] + final["wfh_hours"]).round(2)


final["payable_days"] = (final["total_working_days"] - final["leave_without_paid"]).round(2)

final["effective_working_days_for_hours"] = (final["total_working_days"] - final["leave_taken"]).round(2)
final.loc[final["effective_working_days_for_hours"] < 0, "effective_working_days_for_hours"] = 0.0

final["actual_hours"] = (final["effective_working_days_for_hours"] * HOURS_PER_DAY).round(2)

final["short_hours"] = (final["actual_hours"] - final["top_tracker_hours"]).round(2)
final.loc[final["short_hours"] < 0, "short_hours"] = 0.0

# ✅ NEW: short hours based on biometric attendance
final["short_hours_biometric"] = (final["actual_hours"] - final["biometric_total_hours"]).round(2)
final.loc[final["short_hours_biometric"] < 0, "short_hours_biometric"] = 0.0

# ✅ NEW: total short hours (TopTracker short + Biometric short)
final["total_short_hours"] = (final["short_hours"] + final["short_hours_biometric"]).round(2)

final["avg_screenshot_minutes"] = pd.NA
mask = final["screenshot_count"] > 0
final.loc[mask, "avg_screenshot_minutes"] = (
    (final.loc[mask, "total_tracked_seconds"] / 60.0) / final.loc[mask, "screenshot_count"]
).round(2)

# =========================
# STEP 6: OUTPUT CSV (add new column)
# =========================
final_out = final.rename(columns={
    "employee_code": "Employee id",
    "employee_name": "Name",
    "total_working_days": "Total working days",
    "top_tracker_hours": "Top tracker hours",
    "screenshot_count": "Screen shot count",
    "leave_taken": "Number of leave has taken",
    "leave_without_paid": "Number of leave without paid",
    "work_from_home": "Work from home",
    "payable_days": "Payable days",
    "total_tracked_seconds": "Total tracked seconds",
    "avg_screenshot_minutes": "Avg screenshot minutes",
    "actual_hours": "Actual hours",
    "short_hours": "Short hours",
    "short_hours_biometric": "Short hours by biometric",
    "total_short_hours": "Total short hours",
    "biometric_total_hours": "Biometric total hours",   # ✅ NEW column name
})

final_out = final_out[
    [
        "Employee id",
        "Name",
        "Total working days",
        "Top tracker hours",
        "Biometric total hours",        # ✅ NEW in output
        "Screen shot count",
        "Number of leave has taken",
        "Number of leave without paid",
        "Work from home",
        "Payable days",
        "Total tracked seconds",
        "Avg screenshot minutes",
        "Actual hours",
        "Short hours",
        "Short hours by biometric",
        "Total short hours",
    ]
].sort_values(["Employee id", "Name"]).reset_index(drop=True)

final_out.to_csv(OUTPUT_FINAL, index=False)

# Debug exports
calendar_df.to_csv(OUTPUT_WORKING_CALENDAR, index=False)
toptracker_month.to_csv(OUTPUT_TOPTRACKER_MONTH, index=False)
leave_month.to_csv(OUTPUT_LEAVE_MONTH, index=False)

# ✅ biometric debug exports
biometric_daily.to_csv(OUTPUT_BIOMETRIC_DAILY, index=False)
biometric_month.to_csv(OUTPUT_BIOMETRIC_MONTH, index=False)

print("✅ Generated monthly payroll summary:")
print(" -", OUTPUT_FINAL)
print("Debug files:")
print(" -", OUTPUT_WORKING_CALENDAR)
print(" -", OUTPUT_TOPTRACKER_MONTH)
print(" -", OUTPUT_LEAVE_MONTH)
print(" -", OUTPUT_BIOMETRIC_DAILY)
print(" -", OUTPUT_BIOMETRIC_MONTH)

if os.path.exists(OUTPUT_NAME_MISMATCH):
    print("Name mismatch report:")
    print(" -", OUTPUT_NAME_MISMATCH)
    print("Tip: add confirmed mappings into:")
    print(" -", ALIASES_FILE)

