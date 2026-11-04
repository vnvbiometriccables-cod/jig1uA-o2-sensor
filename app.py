"""
O₂ SENSOR IN ENVIRONMENTAL CHAMBER
Live Sample Maximum / Minimum Monitoring

Reads the live O2_Data tab of the environmental-chamber Google Sheet and reports,
for Sample 36, 28, 27, 30 and 32 (all in µA, never converted):
  • actual maximum / minimum / difference / average over ALL valid numeric readings
    (out-of-range values are included – thresholds never filter the statistics)
  • threshold compliance against an adjustable band (default 50–146 µA, inclusive)
  • live trend graphs, an individual-sample view and a 5-sheet Excel report

Data source
  • "csv" (default): public CSV export – the sheet must be shared
    "Anyone with the link – Viewer" (it is, as of 04 Nov 2026).
  • "api": Google Sheets API with a service account stored in Streamlit secrets.
"""

from __future__ import annotations

import io
import math
import re
import threading
import warnings
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# =============================================================================
# 1. CONFIGURATION
# =============================================================================
APP_TITLE = "O₂ SENSOR IN ENVIRONMENTAL CHAMBER"
APP_SUBTITLE = "Live Sample Maximum / Minimum Monitoring"
TREND_TITLE = f"{APP_TITLE} – LIVE SENSOR OUTPUT TREND"
UNIT = "µA"

DEFAULT_SHEET_ID = "1fJjKezE5Dujuc_QTmtwefe6z2h8ff-E3QuOO49Yjd1g"
DEFAULT_GID = "1371427582"

SAMPLES = ["Sample 36", "Sample 28", "Sample 27", "Sample 30", "Sample 32"]
SAMPLE_NO = {s: s.split()[-1] for s in SAMPLES}
COLOR = {"Sample 36": "#2563EB", "Sample 28": "#0D9488", "Sample 27": "#7C3AED",
         "Sample 30": "#DB2777", "Sample 32": "#475569"}

DEFAULT_MIN = 50.0
DEFAULT_MAX = 146.0

REFRESH_SECONDS = 60
CACHE_TTL = 50                      # seconds; shorter than the refresh interval
REQUEST_TIMEOUT = (10, 90)          # connect, read (the sheet is several MB)
EXCEL_MAX_ROWS = 1_048_576

WITHIN, BELOW, ABOVE, NODATA = "WITHIN RANGE", "BELOW LIMIT", "ABOVE LIMIT", "NO DATA"
STATUS_COLOR = {WITHIN: "#15803D", BELOW: "#B91C1C", ABOVE: "#EA580C", NODATA: "#6B7280"}
STATUS_BG = {WITHIN: "#DCFCE7", BELOW: "#FEE2E2", ABOVE: "#FFEDD5", NODATA: "#E5E7EB"}


def _setting(key: str, default: str) -> str:
    try:
        return str(st.secrets.get("app_settings", {}).get(key, default))
    except Exception:
        return default


def _has_service_account() -> bool:
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


SHEET_ID = _setting("sheet_id", DEFAULT_SHEET_ID)
GID = _setting("gid", DEFAULT_GID)
DATA_SOURCE = _setting("data_source", "csv").lower()          # csv | api
DISPLAY_TZ = ZoneInfo(_setting("display_timezone", "Asia/Kolkata"))
TIMESTAMP_ORDER = _setting("timestamp_order", "auto").lower()  # auto | monthfirst | dayfirst
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid={GID}"


# =============================================================================
# 2. FETCHING (cached briefly; never fabricates data)
# =============================================================================
class SheetAccessError(Exception):
    pass


def _unique_headers(cols) -> list[str]:
    out, seen = [], {}
    for i, c in enumerate(cols):
        name = str(c).strip() or f"Column {i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 1
        out.append(name)
    return out


def fetch_csv() -> pd.DataFrame:
    try:
        r = requests.get(CSV_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True,
                         headers={"User-Agent": "O2-sensor-dashboard/1.0"})
    except requests.Timeout as e:
        raise SheetAccessError("Google Sheets did not respond in time (request timeout).") from e
    except requests.ConnectionError as e:
        raise SheetAccessError("No connection to Google Sheets. Check the internet connection.") from e
    if r.status_code in (401, 403):
        raise SheetAccessError("Google Sheets refused access. Share the sheet as “Anyone with the link – "
                               "Viewer”, or switch to the service-account method (see README).")
    if r.status_code == 404:
        raise SheetAccessError("Sheet or tab not found. Check the sheet ID and gid.")
    if r.status_code >= 400:
        raise SheetAccessError(f"Google Sheets returned HTTP {r.status_code}.")
    ctype = r.headers.get("Content-Type", "")
    text = r.content.decode("utf-8-sig", errors="replace")
    if "text/html" in ctype or text.lstrip()[:15].lower().startswith(("<!doctype", "<html")):
        raise SheetAccessError("Google returned a sign-in page instead of CSV: the sheet is not public. "
                               "Share it as “Anyone with the link – Viewer”, or use the service-account method.")
    if not text.strip():
        raise SheetAccessError("Google Sheets returned an empty CSV.")
    try:
        df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, na_filter=False,
                         skip_blank_lines=True, on_bad_lines="skip")
    except Exception as e:
        raise SheetAccessError(f"The CSV from Google Sheets could not be read ({type(e).__name__}).") from e
    if df.shape[1] < 2:
        raise SheetAccessError("The CSV has fewer than two columns – check the gid points at the data tab.")
    df.columns = _unique_headers(df.columns)
    return df


@st.cache_resource(show_spinner=False)
def _gspread_client():
    import gspread
    return gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))


def fetch_api() -> pd.DataFrame:
    if not _has_service_account():
        raise SheetAccessError("data_source = \"api\" but [gcp_service_account] is missing from secrets.")
    try:
        sh = _gspread_client().open_by_key(SHEET_ID)
        ws = sh.get_worksheet_by_id(int(GID))
        values = ws.get_all_values()
    except Exception as e:
        name = type(e).__name__
        if name == "SpreadsheetNotFound":
            raise SheetAccessError("Spreadsheet not found or not shared with the service-account e-mail.") from e
        if name == "WorksheetNotFound":
            raise SheetAccessError(f"No tab with gid {GID} in the spreadsheet.") from e
        raise SheetAccessError(f"Google Sheets API error: {name}: {e}") from e
    if len(values) < 2:
        raise SheetAccessError("The sheet tab is empty.")
    width = max(len(r) for r in values)
    rows = [r + [""] * (width - len(r)) for r in values]
    df = pd.DataFrame(rows[1:], columns=_unique_headers(rows[0]), dtype=str)
    return df


@st.cache_resource(ttl=CACHE_TTL, show_spinner=False, max_entries=1)
def fetch_sheet() -> tuple[pd.DataFrame, str, str]:
    """Returns (raw table of strings, fetch id / UTC time, source label)."""
    if DATA_SOURCE == "api":
        df, src = fetch_api(), "Google Sheets API"
    else:
        df, src = fetch_csv(), "CSV export"
    df = df.loc[~(df == "").all(axis=1)].reset_index(drop=True)
    return df, datetime.now(timezone.utc).isoformat(), src


@st.cache_resource(show_spinner=False)
def shared_store() -> dict:
    """Last successful load, shared by every viewer and kept across reruns."""
    return {"lock": threading.Lock(), "last_good": None}


# =============================================================================
# 3. PARSING & CALCULATIONS (pure functions)
# =============================================================================
_NUM_CLEAN = re.compile(r"[^\d.\-+eE]")
_DMY = re.compile(r"^\s*(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})")


def to_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\s*(µa|ua|μa)\s*$", "", regex=True, case=False)
    s = s.where(~(s.str.contains(",") & ~s.str.contains(r"\.")), s.str.replace(",", ".", regex=False))
    out = pd.to_numeric(s, errors="coerce")
    return out.where(np.isfinite(out))


def _dayfirst(sample: list[str]) -> bool:
    if TIMESTAMP_ORDER == "dayfirst":
        return True
    if TIMESTAMP_ORDER == "monthfirst":
        return False
    for v in sample:
        m = _DMY.match(v)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > 12:
                return True
            if b > 12:
                return False
    return False   # ambiguous → month/day/year (the logger's format, e.g. 11/4/2026)


def parse_timestamps(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    nonblank = s.ne("") & s.str.contains(r"\d")
    if not nonblank.any():
        return out
    num = pd.to_numeric(s.where(nonblank), errors="coerce")
    serial = num.between(20000, 100000)
    if serial.any():   # Google Sheets serial dates (API path with date cells)
        out.loc[serial] = (pd.Timestamp("1899-12-30") + pd.to_timedelta(num[serial], unit="D")).dt.round("s")
    txt = nonblank & ~serial
    if txt.any():
        first = s[txt].iloc[:2000].tolist() + s[txt].iloc[-500:].tolist()
        md = "%d/%m/%Y" if _dayfirst(first) else "%m/%d/%Y"
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            for fmt in (f"{md} %H:%M:%S", f"{md} %H:%M", md, md.replace("/", "-") + " %H:%M:%S",
                        "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
                todo = txt & out.isna()
                if not todo.any():
                    break
                out.loc[todo] = pd.to_datetime(s[todo], format=fmt, errors="coerce")
            todo = txt & out.isna()
            if todo.any():
                out.loc[todo] = pd.to_datetime(s[todo], dayfirst=md.startswith("%d"), errors="coerce",
                                               format="mixed")
    return out


def guess_mapping(columns: list[str]) -> dict:
    """Best-guess column for the timestamp and each sample (user can override)."""
    norm = {c: re.sub(r"\s+", " ", c.lower()) for c in columns}
    guess = {}
    ts = [c for c in columns if re.search(r"date|time|stamp", norm[c])]
    guess["Timestamp"] = ts[0] if ts else (columns[0] if columns else None)
    for s in SAMPLES:
        n = SAMPLE_NO[s]
        exact = [c for c in columns if re.search(rf"\bsample\s*{n}\b", norm[c])]
        loose = [c for c in columns if re.search(rf"(?<!\d){n}(?!\d)", norm[c]) and c != guess["Timestamp"]]
        guess[s] = (exact or loose or [None])[0]
    return guess


def classify(values: pd.Series, lo: float, hi: float) -> pd.Series:
    st_ = pd.Series(NODATA, index=values.index, dtype=object)
    st_[values.notna() & (values < lo)] = BELOW
    st_[values.notna() & (values > hi)] = ABOVE
    st_[values.notna() & (values >= lo) & (values <= hi)] = WITHIN
    return st_


def build_dataset(raw: pd.DataFrame, mapping: dict) -> dict:
    """Mapped, parsed, chronologically sorted readings. Values are never altered."""
    n = len(raw)
    ts_col = mapping.get("Timestamp")
    ts = parse_timestamps(raw[ts_col]) if ts_col in raw.columns else pd.Series(pd.NaT, index=raw.index)
    data = pd.DataFrame({"Timestamp": ts, "Sheet Row": np.arange(n) + 2})
    original = pd.DataFrame(index=raw.index)
    missing_cols = []
    for s in SAMPLES:
        col = mapping.get(s)
        if col in raw.columns:
            data[s] = to_numeric(raw[col])
            original[s] = raw[col]
        else:
            data[s] = np.nan
            original[s] = ""
            missing_cols.append(s)
    order = data.sort_values(["Timestamp", "Sheet Row"], na_position="first", kind="stable").index
    data, original = data.loc[order].reset_index(drop=True), original.loc[order].reset_index(drop=True)
    return {
        "data": data, "original": original, "rows": n, "missing_cols": missing_cols,
        "invalid_ts": int(data["Timestamp"].isna().sum()),
        "latest_ts": data["Timestamp"].max(),
    }


def sample_stats(ds: dict, lo: float, hi: float) -> pd.DataFrame:
    d = ds["data"]
    rows = []
    for s in SAMPLES:
        v = d[s]
        valid = v.notna()
        n = int(valid.sum())
        if n:
            imax, imin = v.idxmax(), v.idxmin()
            last = v.last_valid_index()
            within = int(((v >= lo) & (v <= hi)).sum())
            below = int((v < lo).sum())
            above = int((v > hi).sum())
            rows.append({
                "Sample": s, "Maximum (µA)": v[imax], "Maximum Timestamp": d.at[imax, "Timestamp"],
                "Minimum (µA)": v[imin], "Minimum Timestamp": d.at[imin, "Timestamp"],
                "Difference (µA)": v[imax] - v[imin], "Average (µA)": float(v.mean()),
                "Total Readings": n, "Missing/Invalid": int((~valid).sum()),
                "Latest (µA)": v[last], "Latest Timestamp": d.at[last, "Timestamp"],
                "Latest Status": classify(pd.Series([v[last]]), lo, hi).iloc[0],
                "Within Range": within, "Below Limit": below, "Above Limit": above,
                "Compliance %": within / n * 100,
            })
        else:
            rows.append({"Sample": s, "Maximum (µA)": np.nan, "Maximum Timestamp": pd.NaT,
                         "Minimum (µA)": np.nan, "Minimum Timestamp": pd.NaT, "Difference (µA)": np.nan,
                         "Average (µA)": np.nan, "Total Readings": 0, "Missing/Invalid": int(len(v)),
                         "Latest (µA)": np.nan, "Latest Timestamp": pd.NaT, "Latest Status": NODATA,
                         "Within Range": 0, "Below Limit": 0, "Above Limit": 0, "Compliance %": np.nan})
    return pd.DataFrame(rows)


def overall_extremes(stats: pd.DataFrame) -> dict:
    ok = stats[stats["Total Readings"] > 0]
    if ok.empty:
        return {"high": np.nan, "high_sample": "—", "high_ts": pd.NaT,
                "low": np.nan, "low_sample": "—", "low_ts": pd.NaT}
    h = ok.loc[ok["Maximum (µA)"].idxmax()]
    l = ok.loc[ok["Minimum (µA)"].idxmin()]
    return {"high": h["Maximum (µA)"], "high_sample": h["Sample"], "high_ts": h["Maximum Timestamp"],
            "low": l["Minimum (µA)"], "low_sample": l["Sample"], "low_ts": l["Minimum Timestamp"]}


def decimate(t: pd.Series, v: pd.Series, max_points: int = 6000):
    """Display-only thinning that keeps each bin's highest and lowest point, so peaks stay visible."""
    m = v.notna() & t.notna()
    t, v = t[m], v[m]
    if len(v) <= max_points:
        return t, v, False
    nb = max_points // 2
    tn = t.astype("int64")
    b = ((tn - tn.min()) * nb // max(1, (tn.max() - tn.min()) + 1)).to_numpy()
    g = pd.Series(v.to_numpy(), index=np.arange(len(v))).groupby(b)
    keep = np.unique(np.concatenate([g.idxmax().to_numpy(), g.idxmin().to_numpy()]))
    return t.iloc[keep], v.iloc[keep], True


# =============================================================================
# 4. EXCEL REPORT (openpyxl, write-only for large raw sheets)
# =============================================================================
def build_excel(ds: dict, stats: pd.DataFrame, ext: dict, lo: float, hi: float, src: str) -> bytes:
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook(write_only=True)
    thin = Side(style="thin", color="9CA3AF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hfont, hfill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1E293B")
    tfont = Font(bold=True, size=14, color="1E293B")
    num_fmt, ts_fmt = "0.0000", "dd-mmm-yyyy hh:mm:ss"
    now = datetime.now(DISPLAY_TZ)

    def cell(ws, value, *, bold=False, fmt=None, header=False, bordered=False, font=None):
        c = WriteOnlyCell(ws, value=value)
        if header:
            c.font, c.fill, c.border = hfont, hfill, border
            c.alignment = Alignment(wrap_text=True, vertical="center")
        elif bordered:
            c.border = border
        if bold:
            c.font = Font(bold=True)
        if font is not None:
            c.font = font
        if fmt:
            c.number_format = fmt
        return c

    def clean(v):
        if v is None:
            return None
        if isinstance(v, float) and not math.isfinite(v):
            return None
        if v is pd.NaT or (isinstance(v, pd.Timestamp) and pd.isna(v)):
            return None
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        if isinstance(v, (np.integer,)):
            return int(v)
        if isinstance(v, (np.floating,)):
            return float(v) if math.isfinite(v) else None
        return v

    def status_rules(ws, rng):
        for status, fg in ((WITHIN, "DCFCE7"), (BELOW, "FEE2E2"), (ABOVE, "FFEDD5"), (NODATA, "E5E7EB")):
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="equal", formula=[f'"{status}"'],
                                fill=PatternFill("solid", fgColor=fg, bgColor=fg),
                                font=Font(bold=True, color=STATUS_COLOR[status].lstrip("#"))))

    def title_rows(ws, title, width):
        ws.append([cell(ws, f"{APP_TITLE} – {APP_SUBTITLE}", font=tfont)])
        ws.append([cell(ws, f"{title} · Unit {UNIT} · Limits {lo:g}–{hi:g} {UNIT} (inclusive) · "
                            f"Generated {now:%d %b %Y %H:%M:%S} ({DISPLAY_TZ.key})", font=Font(italic=True, size=9))])

    def table(ws, title, headers, rows, fmts, widths, status_col=None):
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A4"
        title_rows(ws, title, len(headers))
        ws.append([cell(ws, h, header=True) for h in headers])
        n = 0
        for r in rows:
            ws.append([cell(ws, clean(v), fmt=fmts[i], bordered=True) for i, v in enumerate(r)])
            n += 1
        last_col = get_column_letter(len(headers))
        ws.auto_filter.ref = f"A3:{last_col}{3 + max(n, 1)}"
        if status_col is not None:
            col = get_column_letter(status_col + 1)
            status_rules(ws, f"{col}4:{col}{3 + max(n, 1)}")

    # Sheet 1 – Maximum Minimum Summary
    ws = wb.create_sheet("Maximum Minimum Summary")
    table(ws, "Maximum Minimum Summary",
          ["Sample", "Maximum Output (µA)", "Maximum Timestamp", "Minimum Output (µA)", "Minimum Timestamp",
           "Max–Min Difference (µA)", "Average Output (µA)", "Total Readings"],
          [[r["Sample"], r["Maximum (µA)"], r["Maximum Timestamp"], r["Minimum (µA)"], r["Minimum Timestamp"],
            r["Difference (µA)"], r["Average (µA)"], r["Total Readings"]] for _, r in stats.iterrows()],
          [None, num_fmt, ts_fmt, num_fmt, ts_fmt, num_fmt, num_fmt, "#,##0"],
          [14, 20, 22, 20, 22, 22, 20, 15])

    # Sheet 2 – Overall Summary
    ws = wb.create_sheet("Overall Summary")
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 34, 34
    ws.freeze_panes = "A4"
    title_rows(ws, "Overall Summary", 2)
    ws.append([cell(ws, "Parameter", header=True), cell(ws, "Value", header=True)])
    kv = [
        ("Dashboard title", f"{APP_TITLE} – {APP_SUBTITLE}", None),
        ("Report generated", now.replace(tzinfo=None), ts_fmt),
        ("Data source", f"{src} · {SHEET_URL}", None),
        ("Measurement unit", UNIT, None),
        ("Minimum threshold (µA)", lo, num_fmt),
        ("Maximum threshold (µA)", hi, num_fmt),
        ("Overall highest output (µA)", ext["high"], num_fmt),
        ("Highest sample", ext["high_sample"], None),
        ("Highest timestamp", ext["high_ts"], ts_fmt),
        ("Overall lowest output (µA)", ext["low"], num_fmt),
        ("Lowest sample", ext["low_sample"], None),
        ("Lowest timestamp", ext["low_ts"], ts_fmt),
        ("Total monitored samples", len(SAMPLES), "0"),
        ("Total valid readings", int(stats["Total Readings"].sum()), "#,##0"),
        ("Missing / invalid sample cells", int(stats["Missing/Invalid"].sum()), "#,##0"),
        ("Latest data timestamp", ds["latest_ts"], ts_fmt),
    ]
    for k, v, f in kv:
        ws.append([cell(ws, k, bordered=True, bold=True), cell(ws, clean(v), fmt=f, bordered=True)])

    # Sheet 3 – Threshold Validation Summary
    ws = wb.create_sheet("Threshold Validation Summary")
    tot_n = int(stats["Total Readings"].sum())
    tot_w = int(stats["Within Range"].sum())
    rows = [[r["Sample"], r["Total Readings"], r["Within Range"], r["Below Limit"], r["Above Limit"],
             r["Compliance %"] / 100 if pd.notna(r["Compliance %"]) else None] for _, r in stats.iterrows()]
    rows.append(["ALL SAMPLES", tot_n, tot_w, int(stats["Below Limit"].sum()), int(stats["Above Limit"].sum()),
                 tot_w / tot_n if tot_n else None])
    table(ws, "Threshold Validation Summary",
          ["Sample", "Total Valid Readings", f"Within Range ({lo:g}–{hi:g} µA)", f"Below {lo:g} µA",
           f"Above {hi:g} µA", "Compliance %"],
          rows, [None, "#,##0", "#,##0", "#,##0", "#,##0", "0.00%"], [16, 20, 24, 16, 16, 15])

    d, orig = ds["data"], ds["original"]
    ts_py = [None if pd.isna(t) else t.to_pydatetime() for t in d["Timestamp"]]

    # Sheet 4 – Raw Readings (original values, nothing removed)
    ws = wb.create_sheet("Raw Readings")
    for i, w in enumerate([22] + [14] * len(SAMPLES) + [11], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    title_rows(ws, "Raw Readings – original values, out-of-range readings kept", 7)
    ws.append([cell(ws, h, header=True) for h in ["Timestamp"] + [f"{s} (µA)" for s in SAMPLES] + ["Sheet Row"]])
    vals = {s: d[s].to_numpy() for s in SAMPLES}
    ovals = {s: orig[s].to_numpy() for s in SAMPLES}
    rows_n = len(d)
    for i in range(rows_n):
        row = [ts_py[i]]
        for s in SAMPLES:
            x = vals[s][i]
            row.append(float(x) if x == x else (ovals[s][i] or None))
        row.append(int(d["Sheet Row"].iat[i]))
        ws.append(row)
    ws.auto_filter.ref = f"A3:G{3 + max(rows_n, 1)}"

    # Sheet 5 – Sample-wise Analysis (long format; continues on extra sheets past Excel's row limit)
    per_sheet = EXCEL_MAX_ROWS - 3
    sheet_no, written, ws = 1, 0, None

    def new_long_sheet(k):
        w = wb.create_sheet("Sample-wise Analysis" if k == 1 else f"Sample-wise Analysis ({k})")
        for i, wd in enumerate([22, 14, 20, 16], start=1):
            w.column_dimensions[get_column_letter(i)].width = wd
        w.freeze_panes = "A4"
        title_rows(w, "Sample-wise Analysis" + ("" if k == 1 else f" – continued ({k})"), 4)
        w.append([cell(w, h, header=True) for h in ["Timestamp", "Sample", "Sensor Output (µA)", "Threshold Status"]])
        return w

    def close_long_sheet(w, n):
        w.auto_filter.ref = f"A3:D{3 + max(n, 1)}"
        status_rules(w, f"D4:D{3 + max(n, 1)}")

    ws = new_long_sheet(1)
    for s in SAMPLES:
        v = vals[s]
        idx = np.flatnonzero(~np.isnan(v))
        status = classify(pd.Series(v[idx]), lo, hi).to_numpy()
        for j, i in enumerate(idx):
            if written >= per_sheet:
                close_long_sheet(ws, written)
                sheet_no += 1
                ws, written = new_long_sheet(sheet_no), 0
            ws.append([ts_py[i], s, float(v[i]), status[j]])
            written += 1
    close_long_sheet(ws, written)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@st.cache_data(show_spinner=False, max_entries=2)
def excel_cached(key: str, _ds, _stats, _ext, lo, hi, src) -> bytes:
    return build_excel(_ds, _stats, _ext, lo, hi, src)


# =============================================================================
# 5. UI HELPERS
# =============================================================================
def fmt_num(v, nd=4, unit=f" {UNIT}"):
    try:
        if v is None or pd.isna(v) or not math.isfinite(float(v)):
            return NODATA
    except (TypeError, ValueError):
        return NODATA
    return f"{float(v):,.{nd}f}{unit}"


def fmt_ts(t) -> str:
    return "—" if t is None or pd.isna(t) else pd.Timestamp(t).strftime("%d %b %Y %H:%M:%S")


def fmt_utc(iso: str | None) -> str:
    if not iso:
        return "—"
    return datetime.fromisoformat(iso).astimezone(DISPLAY_TZ).strftime("%d %b %Y %H:%M:%S")


def chip(status: str) -> str:
    return (f'<span class="chip" style="background:{STATUS_BG[status]};color:{STATUS_COLOR[status]}">'
            f'{status}</span>')


def trend_figure(d: pd.DataFrame, samples, lo, hi, title, height=480):
    fig = go.Figure()
    thinned = False
    for s in samples:
        t, v, th = decimate(d["Timestamp"], d[s])
        thinned |= th
        if v.empty:
            continue
        fig.add_trace(go.Scattergl(
            x=t, y=v, mode="lines", name=s, line=dict(color=COLOR[s], width=1.6),
            hovertemplate=f"<b>{s}</b><br>%{{x|%d %b %Y %H:%M:%S}}<br>%{{y:.4f}} {UNIT}<extra></extra>"))
        fig.add_trace(go.Scatter(
            x=[t.iloc[-1]], y=[v.iloc[-1]], mode="markers+text", showlegend=False, hoverinfo="skip",
            marker=dict(size=11, color=COLOR[s], line=dict(width=2, color="white")),
            text=[f"{v.iloc[-1]:.2f}"], textposition="middle right", textfont=dict(color=COLOR[s], size=11)))
    fig.add_hline(y=lo, line_dash="dash", line_color=STATUS_COLOR[BELOW], line_width=1.5,
                  annotation_text=f"Minimum threshold {lo:g} {UNIT}", annotation_position="bottom left",
                  annotation_font_color=STATUS_COLOR[BELOW])
    fig.add_hline(y=hi, line_dash="dash", line_color=STATUS_COLOR[ABOVE], line_width=1.5,
                  annotation_text=f"Maximum threshold {hi:g} {UNIT}", annotation_position="top left",
                  annotation_font_color=STATUS_COLOR[ABOVE])
    fig.update_layout(
        title=dict(text=title, x=0, font=dict(size=15)), height=height, hovermode="closest", dragmode="zoom",
        margin=dict(l=10, r=50, t=55, b=10), legend=dict(orientation="h", yanchor="bottom", y=-0.25, x=0),
        xaxis=dict(title="Timestamp", tickformat="%d %b\n%H:%M"),
        yaxis=dict(title=f"Sensor Current Output ({UNIT})"))
    return fig, thinned


CSS = """
<style>
.block-container {padding-top: 1.5rem;}
.app-title {font-size: 1.65rem; font-weight: 800; letter-spacing: .03em; margin: 0;}
.app-sub {font-size: 1.05rem; font-weight: 600; margin: .1rem 0 .2rem; opacity: .85;}
.status-bar {display:flex; flex-wrap:wrap; gap:.45rem 1.4rem; align-items:center; padding:.65rem 1rem;
  border:1px solid rgba(128,128,128,.3); border-radius:.6rem; margin:.4rem 0 1rem;}
.badge {font-weight:800; padding:.22rem .7rem; border-radius:999px; color:#fff; font-size:.8rem; letter-spacing:.06em;}
.b-live {background:#15803D;} .b-err {background:#B91C1C;} .b-nodata {background:#6B7280;}
.dot {display:inline-block; width:.55rem; height:.55rem; border-radius:50%; background:#fff; margin-right:.35rem;
  animation: pulse 1.6s infinite;}
@keyframes pulse {50%{opacity:.25}}
@media (prefers-reduced-motion: reduce) {.dot{animation:none}}
.si {font-size:.85rem;} .si .k {opacity:.65; margin-right:.3rem;}
.chip {display:inline-block; font-weight:800; font-size:.74rem; padding:.15rem .6rem; border-radius:999px; letter-spacing:.04em;}
.card-head {display:flex; justify-content:space-between; align-items:center; gap:.5rem;}
.card-name {font-weight:800; font-size:1.05rem;}
.card-val {font-size:1.7rem; font-weight:700; font-variant-numeric:tabular-nums; margin:.25rem 0 .1rem;}
.card-meta {font-size:.8rem; opacity:.75; line-height:1.45; font-variant-numeric:tabular-nums;}
.stripe {height:4px; border-radius:2px; margin:-.2rem 0 .5rem;}
</style>
"""

PAGES = ["🟢 Live Overview", "📏 Maximum / Minimum", "📈 Live Trend", "🔎 Individual Sample",
         "✅ Threshold Validation", "🗂️ Data Quality"]
NUM4 = st.column_config.NumberColumn(format="%.4f")
PCT = st.column_config.NumberColumn(format="%.2f %%")
DTC = st.column_config.DatetimeColumn(format="DD MMM YYYY HH:mm:ss")


# =============================================================================
# 6. PAGES
# =============================================================================
def page_overview(ds, stats, ext, lo, hi, status):
    k = st.columns(6)
    k[0].metric("Monitored samples", "5", border=True)
    k[1].metric("Total recorded readings", f"{int(stats['Total Readings'].sum()):,}", border=True,
                help="Valid numeric readings across the five samples.")
    k[2].metric("Overall highest", fmt_num(ext["high"]), border=True,
                help=f"{ext['high_sample']} at {fmt_ts(ext['high_ts'])}")
    k[3].metric("Overall lowest", fmt_num(ext["low"]), border=True,
                help=f"{ext['low_sample']} at {fmt_ts(ext['low_ts'])}")
    k[4].metric("Last successful refresh", fmt_utc(status["fetched_at"]).split(" ", 3)[-1], border=True,
                help=fmt_utc(status["fetched_at"]))
    k[5].metric("Connection", status["label"], border=True)
    st.caption(f"Highest: **{ext['high_sample']}** at {fmt_ts(ext['high_ts'])} · "
               f"Lowest: **{ext['low_sample']}** at {fmt_ts(ext['low_ts'])}")

    cols = st.columns(5)
    for c, (_, r) in zip(cols, stats.iterrows()):
        stt = r["Latest Status"]
        with c, st.container(border=True):
            st.markdown(
                f'<div class="stripe" style="background:{STATUS_COLOR[stt]}"></div>'
                f'<div class="card-head"><span class="card-name">{r["Sample"]}</span>{chip(stt)}</div>'
                f'<div class="card-val">{fmt_num(r["Latest (µA)"])}</div>'
                f'<div class="card-meta">Limits {lo:g} – {hi:g} {UNIT}<br>'
                f'Latest reading {fmt_ts(r["Latest Timestamp"])}<br>'
                f'Total readings {int(r["Total Readings"]):,}<br>'
                f'Refreshed {fmt_utc(status["fetched_at"])}</div>', unsafe_allow_html=True)

    fig, thinned = trend_figure(ds["data"].loc[ds["data"]["Timestamp"] >=
                                               ds["latest_ts"] - pd.Timedelta(hours=24)] if pd.notna(ds["latest_ts"])
                                else ds["data"], SAMPLES, lo, hi, f"{TREND_TITLE} (last 24 hours)", height=400)
    st.plotly_chart(fig, width="stretch", key="ov_trend")


def page_maxmin(stats, ext):
    st.caption("Maximum and minimum use **every valid numeric reading**, including readings outside the "
               "threshold band. Thresholds do not filter these values.")
    t = stats[["Sample", "Maximum (µA)", "Minimum (µA)", "Difference (µA)", "Average (µA)", "Total Readings",
               "Maximum Timestamp", "Minimum Timestamp"]].copy()
    for c in ["Maximum (µA)", "Minimum (µA)", "Difference (µA)", "Average (µA)"]:
        t[c] = t[c].map(lambda v: fmt_num(v, unit=""))
    st.dataframe(t, hide_index=True, width="stretch",
                 column_config={"Maximum Timestamp": DTC, "Minimum Timestamp": DTC,
                                "Total Readings": st.column_config.NumberColumn(format="%d")})
    c = st.columns(2)
    with c[0], st.container(border=True):
        st.markdown("**Overall highest reading**")
        st.markdown(f'<div class="card-val">{fmt_num(ext["high"])}</div>'
                    f'<div class="card-meta">{ext["high_sample"]} · {fmt_ts(ext["high_ts"])}</div>',
                    unsafe_allow_html=True)
    with c[1], st.container(border=True):
        st.markdown("**Overall lowest reading**")
        st.markdown(f'<div class="card-val">{fmt_num(ext["low"])}</div>'
                    f'<div class="card-meta">{ext["low_sample"]} · {fmt_ts(ext["low_ts"])}</div>',
                    unsafe_allow_html=True)
    ok = stats[stats["Total Readings"] > 0]
    if not ok.empty:
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Maximum", x=ok["Sample"], y=ok["Maximum (µA)"], marker_color="#EA580C",
                             text=[f"{v:.2f}" for v in ok["Maximum (µA)"]], textposition="outside"))
        fig.add_trace(go.Bar(name="Average", x=ok["Sample"], y=ok["Average (µA)"], marker_color="#64748B",
                             text=[f"{v:.2f}" for v in ok["Average (µA)"]], textposition="outside"))
        fig.add_trace(go.Bar(name="Minimum", x=ok["Sample"], y=ok["Minimum (µA)"], marker_color="#2563EB",
                             text=[f"{v:.2f}" for v in ok["Minimum (µA)"]], textposition="outside"))
        fig.update_layout(barmode="group", height=420, title=dict(text="Maximum / average / minimum per sample", x=0),
                          yaxis_title=f"Sensor Current Output ({UNIT})", margin=dict(l=10, r=10, t=50, b=10),
                          legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(fig, width="stretch", key="mm_bar")


def page_trend(ds, lo, hi):
    ss = st.session_state
    ss.setdefault("trend_samples", list(SAMPLES))
    c = st.columns([3, 1, 1, 3])
    def _set(v):
        ss["trend_samples"] = v
    with c[1]:
        st.button("Select all", width="stretch", on_click=_set, args=(list(SAMPLES),))
    with c[2]:
        st.button("Deselect all", width="stretch", on_click=_set, args=([],))
    with c[0]:
        sel = st.multiselect("Samples", SAMPLES, key="trend_samples")
    windows = {"Last 1 hour": 1, "Last 6 hours": 6, "Last 24 hours": 24, "Last 3 days": 72,
               "Last 7 days": 168, "All data": None}
    with c[3]:
        win = st.segmented_control("Time window", list(windows), default="Last 24 hours", key="trend_win") \
            or "Last 24 hours"
    d = ds["data"]
    if windows[win] and pd.notna(ds["latest_ts"]):
        d = d[d["Timestamp"] >= ds["latest_ts"] - pd.Timedelta(hours=windows[win])]
    if not sel:
        st.info("No samples selected. Use **Select all** or pick samples above.")
        return
    fig, thinned = trend_figure(d, sel, lo, hi, TREND_TITLE)
    st.plotly_chart(fig, width="stretch", key="trend_fig",
                    config={"scrollZoom": True, "displaylogo": False})
    st.caption("Drag to zoom, double-click to reset, use the toolbar to pan. "
               + ("Long windows are thinned for display while keeping each interval's highest and lowest "
                  "point; all statistics use every reading." if thinned else ""))


def page_individual(ds, stats, lo, hi):
    s = st.selectbox("Sample", SAMPLES, key="ind_sample")
    r = stats.set_index("Sample").loc[s]
    k = st.columns(5)
    k[0].metric("Maximum", fmt_num(r["Maximum (µA)"]), border=True)
    k[1].metric("Minimum", fmt_num(r["Minimum (µA)"]), border=True)
    k[2].metric("Average", fmt_num(r["Average (µA)"]), border=True)
    k[3].metric("Latest", fmt_num(r["Latest (µA)"]), border=True)
    k[4].metric("Max–min difference", fmt_num(r["Difference (µA)"]), border=True)
    k = st.columns(5)
    k[0].metric("Total readings", f"{int(r['Total Readings']):,}", border=True)
    k[1].metric("Within range", f"{int(r['Within Range']):,}", border=True)
    k[2].metric(f"Below {lo:g} µA", f"{int(r['Below Limit']):,}", border=True)
    k[3].metric(f"Above {hi:g} µA", f"{int(r['Above Limit']):,}", border=True)
    k[4].metric("Within range %", NODATA if pd.isna(r["Compliance %"]) else f"{r['Compliance %']:.2f} %",
                border=True)
    fig, _ = trend_figure(ds["data"], [s], lo, hi, f"{s} – sensor output trend", height=420)
    st.plotly_chart(fig, width="stretch", key="ind_fig", config={"scrollZoom": True, "displaylogo": False})
    d = ds["data"]
    v = d[s]
    tbl = pd.DataFrame({"Timestamp": d["Timestamp"], "Sensor Output (µA)": v, "Status": classify(v, lo, hi)})
    tbl = tbl[v.notna()].iloc[::-1]
    st.markdown(f"**Reading table** · {len(tbl):,} readings, newest first. Click a column header to sort.")
    st.dataframe(tbl, hide_index=True, width="stretch", height=420,
                 column_config={"Timestamp": DTC, "Sensor Output (µA)": NUM4})


def page_validation(stats, lo, hi):
    st.caption(f"Each valid numeric reading is compared with the limits: below {lo:g} µA = {BELOW}, "
               f"{lo:g}–{hi:g} µA inclusive = {WITHIN}, above {hi:g} µA = {ABOVE}. Missing or invalid cells "
               f"are counted separately and never as within range.")
    t = stats[["Sample", "Total Readings", "Within Range", "Below Limit", "Above Limit", "Compliance %",
               "Missing/Invalid"]].rename(columns={"Total Readings": "Total Valid Readings",
                                                   "Below Limit": f"Below {lo:g} µA", "Above Limit": f"Above {hi:g} µA",
                                                   "Missing/Invalid": "Missing / Invalid Cells"})
    st.dataframe(t, hide_index=True, width="stretch", column_config={"Compliance %": PCT})
    n, w = int(stats["Total Readings"].sum()), int(stats["Within Range"].sum())
    k = st.columns(4)
    k[0].metric("Total within range", f"{w:,}", border=True)
    k[1].metric("Total below limit", f"{int(stats['Below Limit'].sum()):,}", border=True)
    k[2].metric("Total above limit", f"{int(stats['Above Limit'].sum()):,}", border=True)
    k[3].metric("Overall compliance", f"{w / n * 100:.2f} %" if n else NODATA, border=True)
    ok = stats[stats["Total Readings"] > 0]
    if not ok.empty:
        fig = go.Figure()
        for col, name, colr in (("Within Range", WITHIN, STATUS_COLOR[WITHIN]), ("Below Limit", BELOW, STATUS_COLOR[BELOW]),
                                ("Above Limit", ABOVE, STATUS_COLOR[ABOVE])):
            fig.add_trace(go.Bar(name=name, x=ok["Sample"], y=ok[col] / ok["Total Readings"] * 100, marker_color=colr,
                                 hovertemplate="%{x}<br>%{y:.2f} %<extra>" + name + "</extra>"))
        fig.update_layout(barmode="stack", height=360, yaxis_title="Share of valid readings (%)",
                          title=dict(text="Threshold compliance by sample", x=0), margin=dict(l=10, r=10, t=50, b=10),
                          legend=dict(orientation="h", y=-0.18))
        st.plotly_chart(fig, width="stretch", key="val_bar")


def page_quality(ds, raw_cols, mapping, src):
    k = st.columns(4)
    k[0].metric("Sheet rows read", f"{ds['rows']:,}", border=True)
    k[1].metric("Rows without a valid timestamp", f"{ds['invalid_ts']:,}", border=True,
                help="Their sample readings still count in the statistics; they cannot be placed on the graphs.")
    k[2].metric("Columns in the sheet", f"{len(raw_cols)}", border=True)
    k[3].metric("Data source", src, border=True)
    st.markdown("**Current column mapping**")
    st.dataframe(pd.DataFrame({"Field": ["Timestamp"] + SAMPLES,
                               "Sheet column": [mapping.get(f) or "— not mapped —" for f in ["Timestamp"] + SAMPLES]}),
                 hide_index=True, width="stretch")
    st.markdown("**Latest rows as read from the sheet (mapped columns, original values)**")
    d, o = ds["data"], ds["original"]
    tail = pd.DataFrame({"Timestamp": d["Timestamp"]})
    for s in SAMPLES:
        tail[s] = o[s]
    st.dataframe(tail.tail(200).iloc[::-1], hide_index=True, width="stretch", height=360,
                 column_config={"Timestamp": DTC})
    st.caption(f"Sheet: {SHEET_URL}")


# =============================================================================
# 7. MAIN
# =============================================================================
def main():
    st.set_page_config(page_title=f"{APP_TITLE} – {APP_SUBTITLE}", page_icon="🧪", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    ss = st.session_state
    ss.setdefault("auto_refresh", True)
    ss.setdefault("thr_min", DEFAULT_MIN)
    ss.setdefault("thr_max", DEFAULT_MAX)

    with st.sidebar:
        st.markdown("### O₂ sensor · µA")
        page = st.radio("Navigation", PAGES, key="page", label_visibility="collapsed")
        st.divider()
        if st.button("🔄 REFRESH NOW", width="stretch", type="primary"):
            fetch_sheet.clear()
        st.toggle("Auto-refresh", key="auto_refresh")
        st.caption(f"Refresh interval: {REFRESH_SECONDS} s" + ("" if ss["auto_refresh"] else " (paused)"))

    if ss["auto_refresh"]:
        from streamlit_autorefresh import st_autorefresh
        st_autorefresh(interval=REFRESH_SECONDS * 1000, key="autorefresh")

    # ---- load (keep last good data on failure) ----
    store = shared_store()
    status = {"ok": False, "error": None, "fetched_at": None, "label": "CONNECTION ERROR"}
    raw, src = None, "—"
    try:
        raw, fetched_at, src = fetch_sheet()
        with store["lock"]:
            store["last_good"] = (raw, fetched_at, src)
        status.update(ok=True, fetched_at=fetched_at)
    except Exception as exc:  # noqa: BLE001 – the dashboard must never blank on a failed refresh
        status["error"] = str(exc) if isinstance(exc, SheetAccessError) else f"{type(exc).__name__}: {exc}"
        with store["lock"]:
            last = store["last_good"]
        if last is not None:
            raw, status["fetched_at"], src = last

    # ---- sidebar: column mapping + thresholds ----
    with st.sidebar:
        with st.expander("Column mapping", expanded=False):
            if raw is not None:
                cols = list(raw.columns)
                guess = guess_mapping(cols)
                options = ["— not mapped —"] + cols
                for f in ["Timestamp"] + SAMPLES:
                    key = f"map_{f}"
                    if key not in ss or (ss[key] not in options):
                        ss[key] = guess.get(f) or "— not mapped —"
                    st.selectbox(f"{f} column", options, key=key)
                def _reset_mapping(g=guess):
                    for f_ in ["Timestamp"] + SAMPLES:
                        ss[f"map_{f_}"] = g.get(f_) or "— not mapped —"
                st.button("Reset to detected columns", width="stretch", on_click=_reset_mapping)
            else:
                st.caption("Columns appear after the first successful load.")
        with st.expander("Threshold configuration", expanded=False):
            st.number_input(f"Minimum limit ({UNIT})", key="thr_min", step=1.0, format="%.2f")
            st.number_input(f"Maximum limit ({UNIT})", key="thr_max", step=1.0, format="%.2f")
            def _restore():
                ss["thr_min"], ss["thr_max"] = DEFAULT_MIN, DEFAULT_MAX
            st.button("Restore 50 – 146 µA", width="stretch", on_click=_restore)
            st.caption("Both limits are inclusive. Changing them never alters the readings or the "
                       "maximum / minimum values.")

    lo, hi = float(ss["thr_min"]), float(ss["thr_max"])

    st.markdown(f'<p class="app-title">{APP_TITLE}</p><p class="app-sub">{APP_SUBTITLE}</p>', unsafe_allow_html=True)

    if raw is None:
        st.markdown('<div class="status-bar"><span class="badge b-err">CONNECTION ERROR</span>'
                    '<span class="si"><span class="k">Last successful refresh</span><b>—</b></span></div>',
                    unsafe_allow_html=True)
        st.error(f"**Could not load the Google Sheet.** {status['error']}")
        st.info("No readings are shown until the sheet can be read. The dashboard retries every minute"
                + (" while auto-refresh is on." if ss["auto_refresh"] else "; auto-refresh is currently off."))
        return

    if lo >= hi:
        st.error(f"The minimum limit ({lo:g} µA) must be lower than the maximum limit ({hi:g} µA). "
                 "Fix it under **Threshold configuration**.")
        return

    mapping = {f: (None if ss.get(f"map_{f}") in (None, "— not mapped —") else ss.get(f"map_{f}"))
               for f in ["Timestamp"] + SAMPLES}
    ds = dataset_cached(status["fetched_at"], tuple(sorted(mapping.items(), key=lambda x: x[0])), raw, mapping)
    stats = sample_stats(ds, lo, hi)
    ext = overall_extremes(stats)
    has_data = int(stats["Total Readings"].sum()) > 0
    if status["ok"] and not has_data:
        status["label"] = "NO DATA"
    elif status["ok"]:
        status["label"] = "LIVE"

    # ---- status bar ----
    badge = {"LIVE": '<span class="badge b-live"><span class="dot"></span>LIVE</span>',
             "NO DATA": '<span class="badge b-nodata">NO DATA</span>',
             "CONNECTION ERROR": '<span class="badge b-err">CONNECTION ERROR</span>'}[status["label"]]
    new = track_new(ds, status["fetched_at"])
    st.markdown(
        f'<div class="status-bar">{badge}'
        f'<span class="si"><span class="k">Last successful refresh</span><b>{fmt_utc(status["fetched_at"])}</b></span>'
        f'<span class="si"><span class="k">Latest data timestamp</span><b>{fmt_ts(ds["latest_ts"])}</b></span>'
        f'<span class="si"><span class="k">New rows (last refresh)</span><b>+{new}</b></span>'
        f'<span class="si"><span class="k">Auto-refresh</span><b>{"every " + str(REFRESH_SECONDS) + " s" if ss["auto_refresh"] else "off"}</b></span>'
        f'<span class="si"><span class="k">Limits</span><b>{lo:g}–{hi:g} {UNIT}</b></span></div>',
        unsafe_allow_html=True)
    if not status["ok"]:
        st.error(f"**CONNECTION ERROR – showing the last successfully loaded data.** {status['error']}")
    if ds["missing_cols"]:
        st.warning("Not mapped to a sheet column: **" + ", ".join(ds["missing_cols"]) +
                   "**. Set them under **Column mapping** in the sidebar; these samples show NO DATA until then.")
    if not mapping.get("Timestamp"):
        st.warning("No timestamp column is mapped, so graphs cannot be drawn. Set it under **Column mapping**.")

    # ---- Excel (built only when the button is clicked) ----
    key = f"{status['fetched_at']}|{lo}|{hi}|{sorted(mapping.items())}"
    with st.sidebar:
        st.divider()
        st.download_button("⬇️ DOWNLOAD EXCEL REPORT", lambda: excel_cached(key, ds, stats, ext, lo, hi, src),
                           file_name="O2_Sensor_Environmental_Chamber_Monitoring_Report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch", disabled=not has_data, key="dl_excel")
        st.caption("Large reports take a little while to build after clicking.")

    st.subheader(page.split(" ", 1)[1])
    if page == PAGES[0]:
        page_overview(ds, stats, ext, lo, hi, status)
    elif page == PAGES[1]:
        page_maxmin(stats, ext)
    elif page == PAGES[2]:
        page_trend(ds, lo, hi)
    elif page == PAGES[3]:
        page_individual(ds, stats, lo, hi)
    elif page == PAGES[4]:
        page_validation(stats, lo, hi)
    elif page == PAGES[5]:
        page_quality(ds, list(raw.columns), mapping, src)


@st.cache_resource(show_spinner=False, max_entries=3)
def dataset_cached(fetch_id: str, mapping_key: tuple, _raw, _mapping) -> dict:
    return build_dataset(_raw, _mapping)


def track_new(ds, fetched_at) -> int:
    ss = st.session_state
    if ss.get("seen_fetch") != fetched_at:
        prev = ss.get("seen_rows")
        ss["new_rows"] = max(0, ds["rows"] - prev) if prev is not None else 0
        ss["seen_rows"], ss["seen_fetch"] = ds["rows"], fetched_at
    return ss.get("new_rows", 0)


if __name__ == "__main__":
    main()
