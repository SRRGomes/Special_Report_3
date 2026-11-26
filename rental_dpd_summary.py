"""
Rental Shortfall & 90 DPD Analysis (AF / 3W / 2W) - summart_2.xlsx layout.

AS-AT DATE
  Runs on the 1st of each month. The data month is the previous month (A), and the
  report is "as at the 1st of the run month" because the DPDs as at the last day of A
  are only available after that day.   Run 1 Nov 2026 -> A = Oct 2026, "as at 1st November 2026".
  (The Rental Shortfall Report instead is "as at" the last day of its latest month.)

HOW IT WORKS (main() follows these steps)
  1. Period      A = month before the run month (--as-at-month YYYY-MM to override);
                 activations = the 12 months ending the month before A (Oct 2025 - Sep 2026)
  2. Inputs      Credit Score + RBP decisions, BRANCH DISBURSEMENT OLD for the activation
                 window (else the history file), BDO history (Existing/New), DPD file
  3. Activations facilities disbursed in the window (CONEX_DATE)
  4. Decisions   Credit Score first, then RBP, else Manual
  5. Categories  Product (AF / 3W / 2W), segment, Existing / New (v2 rule: another facility
                 open on the run date (1st), or taken in the 12 months before CONEX_DATE)
  6. LENDING     COLLECTION_FOR_THE_MONTH + DUE_FOR_THE_MONTH summed over the 12 month-ends
                 up to A: parquet in LENDING_data/lending_base first, then csv for the rest
                 (each csv month is saved as parquet for next time)
  7. DPD         DPD_Adj of month A from data/DPD_data (FacNo, clientcode, Month, DPD_Adj)
  8. Metrics     per product / customer type / segment / decision:
                   Rental shortfall  No. = facilities, Value = sum LEASE_AMOUNT,
                                     Rate (Auto Finance) = interest rate weighted by LEASE_AMOUNT,
                                     Collection Ratio = sum collection / sum due
                   90 DPD            No. / Value with DPD_Adj >= 90, Ratio = 90 DPD No. / No.
                   Analysis of 90 DPD  facilities by DPD_Adj 0-29 | 30-59 | 60-89 | 90+
  9. Outputs     PDF (one page per product) + Excel (one sheet per product) in reports/<date>/

    python rental_dpd_summary.py
    python rental_dpd_summary.py --as-at-month 2026-10 --dpd-file "D:/DPD_Oct.xlsx" --no-move
"""
import gc
import csv
import os
import sys
import glob
import argparse
import logging
import time
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font
from openpyxl.styles.borders import Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle

import approved_facility_summary as afs
from approved_facility_summary import (
    BASE_TABLE_STYLE, CONTENT_WIDTH, CS_DECISION, CS_DECISION_MAP, CS_KEY, DATA_DIR,
    DECISION_ORDER, DECISION_TINT, GRID_COLOR, HEADER_BLUE, HISTORY_DAYFIRST, LABEL_FILL,
    PAGE_MARGIN, PAGE_SIZE, PRODUCTS, RBP_DECISION, RBP_DECISION_MAP, RBP_KEY, REPORTS_DIR,
    SEGMENT_ORDER, SUBHDR_BLUE, SUBTOTAL_FILL, TABLE_TITLE_BLUE, TITLE_NAVY, TOTAL_FILL,
    _fill, _key, bdo_file_range, categorize_product, categorize_segment, customer_types,
    find_latest, format_amount, format_count, load_decisions, match_decisions, move_inputs,
    ordinal_date, read_report_excel, require_cols, segment_label, to_date, to_number,
)
import rental_shortfall_summary as rss
from rental_shortfall_summary import (
    BDO_AMOUNT_COL, BDO_KEY, BDO_PREFIX, BDO_START_COL, COLLECTION_COL, DUE_COL,
    LENDING_DIR, LENDING_KEY, MONTHEND_COL, MONTHEND_DAYFIRST, classify_customers_v2,
    customer_rule_text, format_ratio, read_lending_csv,
)
from log_config import setup_logging

log = logging.getLogger("rental_dpd_summary")

# ============================================================
# REPORT NAME  (dedicated: Report Title, Email Subject, output file name)
# ============================================================
DPD_REPORT_TITLE = "Rental Shortfall & 90 DPD Analysis"
DPD_EMAIL_SUBJECT = "Rental Shortfall & 90 DPD Analysis"
DPD_FILE_NAME = "Rental Shortfall & 90 DPD Analysis"

# ============================================================
# CONFIGURATION
# ============================================================
DPD_LENDING_MONTHS = 12                 # LENDING months used for the Collection Ratio
ACTIVATION_MONTHS = 12                  # activation window length
ACTIVATION_GAP_MONTHS = 1               # window ends this many months before the as-at month

DPD_DIR = DATA_DIR / "DPD_data"
DPD_KEY = "FACNO"                       # 'FacNo' after normalising the header
DPD_MONTH_COL = "MONTH"
DPD_VALUE_COL = "DPD_ADJ"               # ' DPD_Adj' (leading space) -> DPD_ADJ
DPD_THRESHOLD = 90                      # 90 DPD = DPD_Adj >= 90
# Analysis of 90 DPD buckets: (label, lower inclusive, upper exclusive)
DPD_BUCKETS = [("< 30 days", None, 30), (">30-<60", 30, 60), (">60-<90", 60, 90), (">90", 90, None)]

# Auto Finance 'Rate' = interest rate weighted by LEASE_AMOUNT, from BRANCH DISBURSEMENT OLD
RATE_COL = os.environ.get("DPD_RATE_COL", "INTEREST_RATE").strip().upper().replace(" ", "_")
RATE_CANDIDATES = [RATE_COL, "INTEREST_RATE", "INT_RATE", "EFFECTIVE_RATE", "FLAT_RATE", "RATE"]
RATE_PRODUCTS = ["Auto Finance"]

# Segment rows always shown (summart_2.xlsx); other segments added only when they have data
DPD_TEMPLATE_SEGMENTS = {
    ("Auto Finance", "Existing Customers"): ["HP/Lease", "Smart Draft", "Enhancement", "Pledge Loan"],
    ("3W", "Existing Customers"): ["HP/Lease", "Enhancement", "Pledge Loan"],
    ("2W", "Existing Customers"): ["HP/Lease", "Enhancement", "Pledge Loan"],
}
DPD_DEFAULT_SEGMENTS = ["HP/Lease"]

# ============================================================
# PERIOD
# ============================================================

def activation_window(as_at_month):
    """(first month, last month) of the activations for an as-at month."""
    last = as_at_month - ACTIVATION_GAP_MONTHS
    return last - (ACTIVATION_MONTHS - 1), last


def window_text(first, last, sep="- "):
    return f"{first.strftime('%B %Y')}{sep}{last.strftime('%B %Y')}"


def default_as_at_month():
    """The month before the run month (a November run -> October, 'as at 1st November')."""
    return pd.Period(datetime.now(), "M") - 1


def dpd_period(as_at_month=None):
    """
    For the orchestrator: (activation start date, activation end date, as-at date,
    as-at month, (first, last) activation months).
    """
    as_at_month = pd.Period(as_at_month, "M") if as_at_month else default_as_at_month()
    first, last = activation_window(as_at_month)
    return (first.start_time.date(), last.end_time.date(), as_at_month.end_time.date(),
            as_at_month, (first, last))


def _range_of(path):
    r = bdo_file_range(path)
    return (r[0].normalize(), r[1].normalize()) if r else None


def find_activation_file(window):
    """Newest BRANCH DISBURSEMENT OLD downloaded for exactly the activation window."""
    want = (window[0].start_time.normalize(), window[1].end_time.normalize())
    files = [f for f in rss._bdo_files() if _range_of(f) == want]
    if not files:
        raise FileNotFoundError(
            f"No {BDO_PREFIX} file for {want[0]:%Y-%m-%d} to {want[1]:%Y-%m-%d} (name must contain "
            f"{want[0]:%Y-%m-%d}_to_{want[1]:%Y-%m-%d}); download it with branch_disb_data.py "
            f"--from-date {want[0]:%Y-%m-%d} --to-date {want[1]:%Y-%m-%d}")
    return max(files, key=os.path.getmtime)


def find_history_file(window):
    """Newest multi-month BRANCH DISBURSEMENT OLD that is not the activation download."""
    want = (window[0].start_time.normalize(), window[1].end_time.normalize())
    files = [f for f in rss._bdo_files() if afs.is_history_bdo(f) and _range_of(f) != want]
    if not files:
        raise FileNotFoundError(f"No {BDO_PREFIX} history file (multi-year download for the "
                                f"Existing/New check)")
    return max(files, key=os.path.getmtime)

# ============================================================
# INPUTS
# ============================================================

def _norm_cols(cols):
    return (pd.Index(cols).astype(str).str.strip().str.upper()
            .str.replace(r"\s+", "_", regex=True))


def find_dpd_file(dpd_dir=DPD_DIR):
    files = [Path(f) for ext in ("xlsx", "xls", "csv")
             for f in glob.glob(glob.escape(str(dpd_dir)) + os.sep + f"*.{ext}")
             if not os.path.basename(f).startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No DPD file (.xlsx / .csv) in {dpd_dir}")
    return max(files, key=os.path.getmtime)


def _parse_month(series):
    """Month column -> Period('M'); handles dates, 'October 2026', 'Oct-26', '2026-10', 202610."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.to_period("M")
    text = series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    yyyymm = text.str.fullmatch(r"\d{6}")
    parsed = pd.to_datetime(text.where(~yyyymm), errors="coerce", format="mixed")
    parsed = parsed.fillna(pd.to_datetime(text.where(yyyymm), format="%Y%m", errors="coerce"))
    return parsed.dt.to_period("M")


def read_dpd_table(path):
    """The DPD sheet as a table: header = the first row containing FacNo."""
    path = Path(path)
    raw = (pd.read_csv(path, header=None, dtype=str, encoding_errors="replace")
           if path.suffix.lower() == ".csv" else pd.read_excel(path, header=None))
    header_row = next((i for i, row in raw.head(30).iterrows()
                       if DPD_KEY in set(_norm_cols(row.dropna()))), None)
    if header_row is None:
        raise ValueError(f"{path.name}: no header row with FacNo")
    df = raw.iloc[header_row + 1:].copy()
    df.columns = _norm_cols(raw.iloc[header_row].fillna(""))
    return df


def dpd_file_with_month(month, dpd_dir=DPD_DIR):
    """Name of the newest DPD file that has rows for `month`, else None (data_request_alert)."""
    files = [Path(f) for ext in ("xlsx", "xls", "csv")
             for f in glob.glob(glob.escape(str(dpd_dir)) + os.sep + f"*.{ext}")
             if not os.path.basename(f).startswith("~$")]
    for f in sorted(files, key=os.path.getmtime, reverse=True):
        try:
            df = read_dpd_table(f)
            if DPD_MONTH_COL in df.columns and month in set(_parse_month(df[DPD_MONTH_COL])):
                return f.name
        except Exception as e:
            log.warning(f"DPD {f.name}: cannot check ({e})")
    return None


def load_dpd(path, as_at_month):
    """DPD_Adj per facility for the as-at month (max if a facility appears twice)."""
    path = Path(path)
    df = read_dpd_table(path)
    require_cols(df, [DPD_KEY, DPD_VALUE_COL], f"DPD {path.name}")
    df["_KEY"] = _key(df[DPD_KEY])
    df = df[df["_KEY"].notna()].copy()
    df["_DPD"] = pd.to_numeric(df[DPD_VALUE_COL].astype(str).str.replace(",", ""), errors="coerce")

    if DPD_MONTH_COL in df.columns:
        df["_DPD_MONTH"] = _parse_month(df[DPD_MONTH_COL])
        months = sorted(df["_DPD_MONTH"].dropna().unique())
        log.info(f"DPD {path.name}: {len(df):,} rows | month(s) {[str(m) for m in months]}")
        if as_at_month in months:
            df = df[df["_DPD_MONTH"] == as_at_month]
        elif months:
            log.warning(f"DPD {path.name}: no rows for the as-at month {as_at_month}; "
                        f"using its latest month {months[-1]}")
            df = df[df["_DPD_MONTH"] == months[-1]]
        else:
            log.warning(f"DPD {path.name}: {DPD_MONTH_COL} unreadable; using all rows")
    else:
        log.warning(f"DPD {path.name}: no {DPD_MONTH_COL} column; using all rows")

    bad = int(df["_DPD"].isna().sum())
    if bad:
        log.warning(f"DPD {path.name}: {bad:,} row(s) with a non-numeric {DPD_VALUE_COL} (ignored)")
    df = df.dropna(subset=["_DPD"])
    dup = df["_KEY"].duplicated(keep=False)
    if dup.any():
        log.warning(f"DPD {path.name}: {int(dup.sum()):,} rows share a FacNo; the highest DPD_Adj is used")
    out = df.groupby("_KEY", as_index=False)["_DPD"].max()
    log.info(f"DPD: {len(out):,} facilities | DPD_Adj >= {DPD_THRESHOLD}: "
             f"{int((out['_DPD'] >= DPD_THRESHOLD).sum()):,}")
    return out


def load_activations(activation_file, first, last):
    """Facilities in the activation download (CONEX_DATE first..last)."""
    d = read_report_excel(activation_file, BDO_KEY)
    require_cols(d, [BDO_AMOUNT_COL, "PRODUCT_NAME", "ASSET_DESCR", BDO_START_COL],
                 f"{BDO_PREFIX} {Path(activation_file).name}")
    if "SCHEME_DESCRIPTION" not in d.columns:
        log.warning(f"{Path(activation_file).name}: SCHEME_DESCRIPTION not found; segments use PRODUCT_NAME only")
        d["SCHEME_DESCRIPTION"] = pd.NA
    d[BDO_AMOUNT_COL] = to_number(d[BDO_AMOUNT_COL])
    d[BDO_START_COL] = to_date(d[BDO_START_COL], HISTORY_DAYFIRST)
    r = bdo_file_range(activation_file)
    if r and (r[0] > first.start_time or r[1] < last.end_time.normalize()):
        log.warning(f"{Path(activation_file).name} covers {r[0]:%d-%b-%Y} to {r[1]:%d-%b-%Y}; "
                    f"the activation window is {first} to {last}: activations will be missing")
    month = d[BDO_START_COL].dt.to_period("M")
    outside = month.notna() & ((month < first) | (month > last))
    if outside.any():
        log.warning(f"{Path(activation_file).name}: {int(outside.sum()):,} facility(ies) with "
                    f"{BDO_START_COL} outside {first} to {last}; excluded")
    no_date = month.isna()
    if no_date.any():
        log.warning(f"{Path(activation_file).name}: {int(no_date.sum()):,} facility(ies) without a "
                    f"{BDO_START_COL}; kept")
    d = d[~outside].copy()
    d["_KEY"] = _key(d[BDO_KEY])
    dup = d["_KEY"].duplicated(keep="first")
    if dup.any():
        log.warning(f"{int(dup.sum()):,} duplicate {BDO_KEY}(s) in the activation window; keeping the first")
        d = d[~dup].copy()
    d["_MONTH"] = d[BDO_START_COL].dt.to_period("M")
    d[BDO_START_COL] = d[BDO_START_COL].fillna(first.start_time)
    d["ACTIVATION_MONTH"] = d["_MONTH"].astype(str)
    log.info(f"Activations {first} to {last}: {len(d):,} facilities | per month "
             f"{d['ACTIVATION_MONTH'].value_counts().sort_index().to_dict()}")
    return d


def attach_rate(df):
    """_RATE per facility from the BRANCH DISBURSEMENT OLD rate column."""
    col = next((c for c in RATE_CANDIDATES if c in df.columns), None)
    if col:
        df["_RATE"] = pd.to_numeric(df[col].astype(str).str.replace(",", "").str.rstrip("%"),
                                    errors="coerce")
        log.info(f"Rate: {col} from {BDO_PREFIX} ({int(df['_RATE'].notna().sum()):,} facilities)")
        return df, col
    hints = [c for c in df.columns if "RATE" in c or "INT" in c]
    log.warning(f"Rate column not found in {BDO_PREFIX} (tried {RATE_CANDIDATES}); the Rate "
                f"column is left blank. Set DPD_RATE_COL in credentials.env. Columns that look "
                f"related: {hints}")
    df["_RATE"] = float("nan")
    return df, None


# ============================================================
# LENDING (parquet first, then csv)
# ============================================================

LENDING_BASE_SUBDIR = "lending_base"     # parquet store inside LENDING_data
LENDING_NEEDED_COLS = [MONTHEND_COL, LENDING_KEY, COLLECTION_COL, DUE_COL]
CSV_PEEK_ROWS = 200                       # data rows read to see a csv's month(s)


def needed_lending_months(as_at_month):
    return list(pd.period_range(as_at_month - (DPD_LENDING_MONTHS - 1), as_at_month, freq="M"))


def _norm(name):
    return " ".join(str(name).split()).upper().replace(" ", "_")


def _months_of(series):
    return to_date(series, dayfirst=MONTHEND_DAYFIRST).dt.to_period("M")


def _lending_frame(df, source):
    """Standard shape: _MONTH, _KEY, COLLECTION, DUE (rows without month/key dropped)."""
    out = pd.DataFrame({"_MONTH": _months_of(df[MONTHEND_COL]), "_KEY": _key(df[LENDING_KEY]),
                        COLLECTION_COL: to_number(df[COLLECTION_COL]),
                        DUE_COL: to_number(df[DUE_COL])})
    bad = int(out["_MONTH"].isna().sum())
    if bad:
        log.warning(f"LENDING {source}: {bad:,} row(s) with an unreadable {MONTHEND_COL} (ignored)")
    return out.dropna(subset=["_MONTH", "_KEY"])


def _read_parquet_months(path, wanted):
    """
    Read MONTHEND only; if it holds any wanted month, read the 4 needed columns and keep
    those months' rows. Returns (frame or None, months found in the file).
    """
    import pyarrow.parquet as pq
    names = pq.read_schema(path).names
    actual = {_norm(n): n for n in names}
    missing = [c for c in LENDING_NEEDED_COLS if c not in actual]
    if missing:
        log.warning(f"LENDING parquet {path.name}: column(s) {missing} not found "
                    f"(has {names[:15]}); skipped")
        return None, []
    me = pd.read_parquet(path, columns=[actual[MONTHEND_COL]])
    months = _months_of(me[actual[MONTHEND_COL]])
    found = sorted(months.dropna().unique())
    keep = [m for m in found if m in wanted]
    del me, months
    gc.collect()
    if not keep:
        log.info(f"LENDING parquet {path.name}: month(s) {[str(m) for m in found]} not needed; skipped")
        return None, found
    raw = pd.read_parquet(path, columns=[actual[c] for c in LENDING_NEEDED_COLS])
    raw.columns = LENDING_NEEDED_COLS
    df = _lending_frame(raw, path.name)
    del raw
    df = df[df["_MONTH"].isin(keep)]
    gc.collect()
    log.info(f"LENDING parquet {path.name}: {len(df):,} rows for {[str(m) for m in keep]}"
             + (f" (other months skipped: {[str(m) for m in found if m not in keep]})"
                if len(keep) < len(found) else ""))
    return df, found


def _peek_csv_months(path):
    """Months in the first CSV_PEEK_ROWS data rows (header = row with MONTHEND and NEW_CONNO)."""
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, newline="", encoding=enc) as fh:
                reader = csv.reader(fh)
                idx = None
                for i, row in enumerate(reader):
                    cells = [_norm(c) for c in row]
                    if MONTHEND_COL in cells and LENDING_KEY in cells:
                        idx = cells.index(MONTHEND_COL)
                        break
                    if i > 50:
                        return None
                if idx is None:
                    return None
                vals = []
                for row in reader:
                    if len(row) > idx and row[idx].strip():
                        vals.append(row[idx])
                    if len(vals) >= CSV_PEEK_ROWS:
                        break
            return set(_months_of(pd.Series(vals, dtype=object)).dropna().unique())
        except UnicodeDecodeError:
            continue
    return None


def _save_month_parquet(df, base_dir, month):
    """Cache one csv month as lending_base/LENDING_<YYYY-MM>.parquet."""
    path = base_dir / f"LENDING_{month}.parquet"
    try:
        base_dir.mkdir(parents=True, exist_ok=True)
        out = df[df["_MONTH"] == month]
        pd.DataFrame({MONTHEND_COL: [month.end_time.normalize()] * len(out),
                      LENDING_KEY: out["_KEY"].values,
                      COLLECTION_COL: out[COLLECTION_COL].values,
                      DUE_COL: out[DUE_COL].values}).to_parquet(path, index=False)
        log.info(f"LENDING {month}: saved {len(out):,} rows to {base_dir.name}/{path.name}")
    except Exception as e:
        log.warning(f"Could not save {path.name} ({e}); next run reads the csv again")


def lending_months_on_hand(lending_dir, include_parquet=True):
    """
    {month: file name} of the LENDING months already available (data_request_alert).
    Parquet: MONTHEND column only. csv: header + first rows only (fast).
    """
    lending_dir = Path(lending_dir)
    found = {}
    if include_parquet:
        import pyarrow.parquet as pq
        for path in sorted((lending_dir / LENDING_BASE_SUBDIR).glob("*.parquet")):
            try:
                names = {_norm(n): n for n in pq.read_schema(path).names}
                col = names.get(MONTHEND_COL)
                if col:
                    for m in _months_of(pd.read_parquet(path, columns=[col])[col]).dropna().unique():
                        found.setdefault(m, f"{LENDING_BASE_SUBDIR}/{path.name}")
            except Exception as e:
                log.warning(f"LENDING parquet {path.name}: cannot check ({e})")
    for path in sorted(lending_dir.glob("*.csv")):
        for m in _peek_csv_months(path) or ():
            found.setdefault(m, path.name)
    return found


def load_lending_dpd(lending_dir, as_at_month):
    """
    COLLECTION / DUE per (_MONTH, _KEY) for the 12 months up to as_at_month:
    lending_base/*.parquet first (newest file wins a month), then *.csv for the rest.
    Returns (frame, months found).
    """
    lending_dir = Path(lending_dir)
    base_dir = lending_dir / LENDING_BASE_SUBDIR
    wanted = set(needed_lending_months(as_at_month))
    have, frames = set(), []

    parquets = sorted(base_dir.glob("*.parquet"), key=os.path.getmtime, reverse=True)
    log.info(f"LENDING: need {sorted(str(m) for m in wanted)} | {len(parquets)} parquet file(s) "
             f"in {base_dir}")
    for path in parquets:
        if have >= wanted:
            break
        try:
            df, _ = _read_parquet_months(path, wanted - have)
        except Exception as e:
            log.warning(f"LENDING parquet {path.name}: cannot read ({e}); skipped")
            continue
        if df is not None:
            got = set(df["_MONTH"].unique())
            have |= got
            frames.append(df)

    remaining = wanted - have
    csvs = sorted(lending_dir.glob("*.csv"), key=os.path.getmtime, reverse=True)
    if remaining:
        log.info(f"LENDING: {sorted(str(m) for m in remaining)} not in parquet; checking "
                 f"{len(csvs)} csv file(s) in {lending_dir}")
    for path in csvs:
        if not remaining:
            break
        peek = _peek_csv_months(path)
        if peek is not None and len(peek) == 1 and not (peek & remaining):
            why = "already loaded" if peek & wanted else "not needed"
            log.info(f"LENDING csv {path.name}: month {[str(m) for m in peek]} {why}; skipped")
            continue
        d = read_lending_csv(path)                    # full read (slow)
        df = _lending_frame(d, path.name)
        del d
        found = set(df["_MONTH"].unique())
        df = df[df["_MONTH"].isin(remaining)]
        gc.collect()
        if df.empty:
            log.info(f"LENDING csv {path.name}: month(s) {sorted(str(m) for m in found)} not "
                     f"needed; discarded")
            del df
            gc.collect()
            continue
        got = set(df["_MONTH"].unique())
        log.info(f"LENDING csv {path.name}: {len(df):,} rows for {sorted(str(m) for m in got)}")
        for m in sorted(got):
            _save_month_parquet(df, base_dir, m)
        remaining -= got
        have |= got
        frames.append(df)

    if not frames:
        raise FileNotFoundError(f"No LENDING data for {sorted(str(m) for m in wanted)} in "
                                f"{base_dir} or {lending_dir}")
    df = pd.concat(frames, ignore_index=True)
    del frames
    gc.collect()
    dup = df.duplicated(["_MONTH", "_KEY"], keep=False)
    if dup.any():
        log.warning(f"LENDING: {int(dup.sum()):,} rows share a {LENDING_KEY} within a month; "
                    f"their {COLLECTION_COL} / {DUE_COL} are summed")
    agg = df.groupby(["_MONTH", "_KEY"], as_index=False)[[COLLECTION_COL, DUE_COL]].sum()
    months = sorted(have)
    if remaining:
        log.warning(f"LENDING: no data for {sorted(str(m) for m in remaining)} (parquet or csv)")
    return agg, months

# ============================================================
# PREPARATION
# ============================================================

def prepare(files, lending_dir, as_at_month):
    """Steps 3-7: one row per activated facility with every column the tables need."""
    first, last = activation_window(as_at_month)
    as_at = as_at_month.end_time.normalize()

    # 3. Activations in the window
    df = load_activations(files["ACTIVATIONS"], first, last)

    # 4. Credit decision: Credit Score first, then RBP, else Manual
    cs = load_decisions(files["CREDIT_SCORE"], CS_KEY, CS_DECISION, CS_DECISION_MAP,
                        "Credit Score Decision Report")
    rbp = load_decisions(files["RBP"], RBP_KEY, RBP_DECISION, RBP_DECISION_MAP,
                         "RBP Security Approved")
    df = match_decisions(df, cs, rbp)

    # 5. Product category / segment and Existing / New customer
    df["PRODUCT_CATEGORY"] = [categorize_product(p, a) for p, a in
                              zip(df["PRODUCT_NAME"], df["ASSET_DESCR"])]
    df["PRODUCT_SEGMENT"] = [categorize_segment(p, s) for p, s in
                             zip(df["PRODUCT_NAME"], df["SCHEME_DESCRIPTION"])]
    log.info(f"Product Category: {df['PRODUCT_CATEGORY'].value_counts().to_dict()}")
    run_date = as_at + pd.Timedelta(days=1)                # 1st of the run month
    df = classify_customers_v2(df, files["HISTORY"], run_date)

    # 6. LENDING: collection / due summed over the 12 month-ends
    lending, lmonths = load_lending_dpd(lending_dir, as_at_month)
    log.info(f"LENDING months used for collection / due: {[str(m) for m in lmonths]}")
    if lmonths[-1] != as_at_month:
        log.warning(f"Latest LENDING month {lmonths[-1]} is not the as-at month {as_at_month}")
    if len(lmonths) < DPD_LENDING_MONTHS:
        log.warning(f"Only {len(lmonths)} LENDING month(s) up to {as_at_month} "
                    f"(expected {DPD_LENDING_MONTHS})")
    per_fac = lending.groupby("_KEY").agg(**{COLLECTION_COL: (COLLECTION_COL, "sum"),
                                             DUE_COL: (DUE_COL, "sum"),
                                             "LENDING_MONTHS": ("_MONTH", "nunique")})
    df = df.merge(per_fac, how="left", left_on="_KEY", right_index=True)
    found = df[DUE_COL].notna()
    log.info(f"LENDING {lmonths[0]} to {lmonths[-1]}: {int(found.sum()):,} of {len(df):,} "
             f"activations found")
    if (~found).any():
        log.warning(f"{int((~found).sum()):,} activation(s) not in any LENDING file (counted in "
                    f"No./Value, not in the Collection Ratio). Sample: "
                    f"{df.loc[~found, BDO_KEY].astype(str).tolist()[:10]}")
    df["FOUND_IN_PORTFOLIO"] = found

    # 7. DPD of the as-at month, plus the Auto Finance rate
    dpd = load_dpd(files["DPD"], as_at_month)
    df = df.merge(dpd, how="left", on="_KEY")
    df = df.rename(columns={"_DPD": "DPD_ADJ"})
    has = df["DPD_ADJ"].notna()
    log.info(f"DPD: {int(has.sum()):,} of {len(df):,} activations found in the DPD file")
    if (~has).any():
        log.warning(f"{int((~has).sum()):,} activation(s) not in the DPD file (not in the 90 DPD "
                    f"or bucket counts). Sample: {df.loc[~has, BDO_KEY].astype(str).tolist()[:10]}")
    df["IS_90_DPD"] = df["DPD_ADJ"] >= DPD_THRESHOLD
    df["DPD_BUCKET"] = [dpd_bucket(v) for v in df["DPD_ADJ"]]

    df, rate_source = attach_rate(df)
    df["RATE"] = df.pop("_RATE")

    target = df[df["PRODUCT_CATEGORY"].isin(PRODUCTS)].copy()
    log.info(f"Product Segment (AF/3W/2W): "
             f"{target.groupby(['PRODUCT_CATEGORY', 'PRODUCT_SEGMENT']).size().to_dict()}")
    return df, target, lmonths, rate_source


def dpd_bucket(v):
    if pd.isna(v):
        return ""
    for label, lo, hi in DPD_BUCKETS:
        if (lo is None or v >= lo) and (hi is None or v < hi):
            return label
    return ""

# ============================================================
# SUMMARY ROWS
# ============================================================

def metrics(d, with_rate):
    """Values for one row, in column order."""
    no = len(d)
    value = float(d[BDO_AMOUNT_COL].sum())
    due = d[DUE_COL].sum()
    coll_ratio = d[COLLECTION_COL].sum() / due if due else float("nan")
    r = d.dropna(subset=["RATE"])
    w = r[BDO_AMOUNT_COL].sum()
    rate = (r["RATE"] * r[BDO_AMOUNT_COL]).sum() / w if w else float("nan")
    dpd = d[d["IS_90_DPD"]]
    dpd_ratio = len(dpd) / no if no else float("nan")
    buckets = [int((d["DPD_BUCKET"] == label).sum()) for label, _, _ in DPD_BUCKETS]
    out = [("count", no), ("amount", value)]
    if with_rate:
        out.append(("rate", rate))
    out += [("ratio", coll_ratio), ("count", len(dpd)), ("amount", float(dpd[BDO_AMOUNT_COL].sum())),
            ("ratio", dpd_ratio)] + [("count", b) for b in buckets]
    return out


def column_groups(with_rate):
    """[(group title, [column titles])] after the 3 label columns."""
    shortfall = ["No.", "Value"] + (["Rate"] if with_rate else []) + ["Collection Ratio"]
    return [("Rental shortfall", shortfall), ("90 DPD", ["No.", "Value", "Ratio"]),
            ("Analysis of 90 DPD", [b[0] for b in DPD_BUCKETS])]


def dpd_segments(product, ctype, data):
    base = DPD_TEMPLATE_SEGMENTS.get((product, ctype), DPD_DEFAULT_SEGMENTS)
    present = set(data["PRODUCT_SEGMENT"])
    return [s for s in SEGMENT_ORDER if s in base or s in present]


def dpd_rows(product, pdata):
    with_rate = product in RATE_PRODUCTS
    rows = []
    for ctype in customer_types(pdata):
        cdata = pdata[pdata["CUSTOMER_TYPE"] == ctype]
        rows.append(("type", ctype, "", None))
        for seg in dpd_segments(product, ctype, cdata):
            sdata = cdata[cdata["PRODUCT_SEGMENT"] == seg]
            label = segment_label(seg, ctype)
            for i, dec in enumerate(DECISION_ORDER):
                rows.append(("decision", label if i == 0 else "", dec,
                             metrics(sdata[sdata["DECISION"] == dec], with_rate)))
            rows.append(("subtotal", f"Total - {label}", "", metrics(sdata, with_rate)))
        rows.append(("type_total", f"Total - {ctype}", "", metrics(cdata, with_rate)))
    rows.append(("grand", "Total", "", metrics(pdata, with_rate)))
    return rows


def fmt(kind, v):
    if kind == "count":
        return format_count(v)
    if kind == "amount":
        return format_amount(v)
    if kind == "rate":
        return "-" if pd.isna(v) else f"{v:.2f}%"
    return format_ratio(v)


EXCEL_FORMATS = {"count": '#,##0;-#,##0;"-"', "amount": '#,##0;-#,##0;"-"',
                 "rate": '0.00"%"', "ratio": '0.0%'}

# ============================================================
# EXCEL (summart_2.xlsx layout)
# ============================================================

def create_excel(target, xlsx_path, as_at, window):
    thin = Side(style="thin", color="2F5496")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill, sub_fill = _fill("#4472C4"), _fill("#8EAADB")
    white = Font(bold=True, color="FFFFFF")
    fills = {"type": _fill("#EAF0FA"), "subtotal": _fill("#DEEAF6"),
             "type_total": _fill("#BDD7EE"), "grand": _fill("#BDD7EE")}
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wb = Workbook()
    wb.remove(wb.active)
    for product in PRODUCTS:
        ws = wb.create_sheet(product)
        pdata = target[target["PRODUCT_CATEGORY"] == product]
        groups = column_groups(product in RATE_PRODUCTS)
        # data column positions, with one blank gap column between groups (template)
        positions, col = [], 5
        for gi, (_, cols) in enumerate(groups):
            positions.append(list(range(col, col + len(cols))))
            col += len(cols) + 1
        data_cols = [c for g in positions for c in g]
        last_col = data_cols[-1]

        ws.cell(row=1, column=2, value=f"{product} (as at {ordinal_date(as_at)})").font = \
            Font(bold=True, size=14, color="1F3864")
        ws.cell(row=2, column=2, value=f"{window_text(*window)} (Activation)").font = \
            Font(italic=True, color="2E4C7E")
        ws.cell(row=3, column=2, value=customer_rule_text(as_at)).font = Font(italic=True, size=9,
                                                                             color="808080")
        for col_, text in zip((2, 3, 4), ("Customer Type", "Product Segment", "Decision")):
            ws.cell(row=4, column=col_, value=text)
            ws.merge_cells(start_row=4, start_column=col_, end_row=5, end_column=col_)
        for (title, cols), pos in zip(groups, positions):
            ws.cell(row=4, column=pos[0], value=title)
            ws.merge_cells(start_row=4, start_column=pos[0], end_row=4, end_column=pos[-1])
            for c, text in zip(pos, cols):
                ws.cell(row=5, column=c, value=text)
        for r in (4, 5):
            for c in list(range(2, 5)) + data_cols:
                cell = ws.cell(row=r, column=c)
                cell.fill = hdr_fill if (r == 4 or c < 5) else sub_fill
                cell.font, cell.alignment, cell.border = white, center, border

        r, seg_start = 6, None
        for kind, label, dec, vals in dpd_rows(product, pdata):
            label_cells = [ws.cell(row=r, column=c) for c in (2, 3, 4)]
            value_cells = [ws.cell(row=r, column=c) for c in data_cols]
            for c in label_cells + value_cells:
                c.border = border
            if kind == "type":
                label_cells[0].value = label
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
            elif kind == "decision":
                if label:
                    label_cells[1].value = label
                    seg_start = r
                label_cells[2].value = dec
                label_cells[2].fill = _fill(DECISION_TINT[dec])
                if dec == DECISION_ORDER[-1]:
                    ws.merge_cells(start_row=seg_start, start_column=3, end_row=r, end_column=3)
                    ws.cell(row=seg_start, column=3).alignment = Alignment(vertical="center",
                                                                          wrap_text=True)
            elif kind == "subtotal":
                label_cells[1].value = label
                ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
            else:
                label_cells[0].value = label
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
            if vals is not None:
                for cell, (k, v) in zip(value_cells, vals):
                    if k == "ratio":
                        cell.value = None if pd.isna(v) else round(v, 4)
                    elif k == "rate":
                        cell.value = None if pd.isna(v) else round(v, 4)
                    else:
                        cell.value = round(v, 2) if k == "amount" else v
                    cell.number_format = EXCEL_FORMATS[k]
            if kind in fills:
                start = 1 if kind == "subtotal" else 0
                for c in (label_cells[start:] + value_cells):
                    c.fill = fills[kind]
                    c.font = Font(bold=True)
            r += 1

        ws.column_dimensions["A"].width = 2
        for col_, width in zip("BCD", (22, 26, 11)):
            ws.column_dimensions[col_].width = width
        for (_, cols), pos in zip(groups, positions):
            for c, text in zip(pos, cols):
                ws.column_dimensions[get_column_letter(c)].width = 18 if text == "Value" else 11
            ws.column_dimensions[get_column_letter(pos[-1] + 1)].width = 2
        ws.freeze_panes = "E6"

    wb.save(xlsx_path)
    log.info(f"Excel saved: {xlsx_path}")
    return xlsx_path

# ============================================================
# PDF (one page per product)
# ============================================================

def _product_table(product, pdata):
    groups = column_groups(product in RATE_PRODUCTS)
    titles = [t for _, cols in groups for t in cols]
    n = len(titles)
    hdr0 = ["Customer Type", "Product Segment", "Decision"]
    for title, cols in groups:
        hdr0 += [title] + [""] * (len(cols) - 1)
    hdr1 = ["", "", ""] + [t.replace("Collection Ratio", "Collection\nRatio")
                           .replace("Value", "Value (LKR)") for t in titles]
    data = [hdr0, hdr1]
    style = list(BASE_TABLE_STYLE) + [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BLUE),
        ('BACKGROUND', (0, 1), (2, 1), HEADER_BLUE),
        ('BACKGROUND', (3, 1), (-1, 1), SUBHDR_BLUE),
        ('TEXTCOLOR', (0, 0), (-1, 1), colors.white),
        ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 1), 'MIDDLE'),
        ('FONTSIZE', (0, 1), (-1, 1), 7),
        ('ALIGN', (3, 2), (-1, -1), 'RIGHT'),
        ('ALIGN', (2, 2), (2, -1), 'CENTER'),
        ('FONTSIZE', (0, 2), (-1, -1), 7),
    ]
    for col in range(3):
        style.append(('SPAN', (col, 0), (col, 1)))
    c0 = 3
    for _, cols in groups:
        style += [('SPAN', (c0, 0), (c0 + len(cols) - 1, 0)),
                  ('LINEAFTER', (c0 + len(cols) - 1, 0), (c0 + len(cols) - 1, -1), 1.0, GRID_COLOR)]
        c0 += len(cols)

    seg_start = None
    for kind, label, dec, vals in dpd_rows(product, pdata):
        r = len(data)
        values = [""] * n if vals is None else [fmt(k, v) for k, v in vals]
        if kind == "type":
            data.append([label, "", ""] + values)
            style += [('SPAN', (0, r), (-1, r)), ('BACKGROUND', (0, r), (-1, r), LABEL_FILL),
                      ('FONTNAME', (0, r), (-1, r), 'Helvetica-Bold'),
                      ('TEXTCOLOR', (0, r), (-1, r), TABLE_TITLE_BLUE)]
        elif kind == "decision":
            data.append(["", label, dec] + values)
            if label:
                seg_start = r
            style.append(('BACKGROUND', (2, r), (2, r), colors.HexColor(DECISION_TINT[dec])))
            if dec == DECISION_ORDER[-1]:
                style += [('SPAN', (1, seg_start), (1, r)), ('SPAN', (0, seg_start), (0, r))]
        elif kind == "subtotal":
            data.append(["", label, ""] + values)
            style += [('SPAN', (1, r), (2, r)), ('BACKGROUND', (1, r), (-1, r), SUBTOTAL_FILL),
                      ('FONTNAME', (1, r), (-1, r), 'Helvetica-Bold')]
        else:
            data.append([label, "", ""] + values)
            style += [('SPAN', (0, r), (2, r)), ('BACKGROUND', (0, r), (-1, r), TOTAL_FILL),
                      ('FONTNAME', (0, r), (-1, r), 'Helvetica-Bold')]
            if kind == "grand":
                style.append(('LINEABOVE', (0, r), (-1, r), 1.2, GRID_COLOR))

    label_w = [1.15 * inch, 1.45 * inch, 0.6 * inch]
    weights = [1.6 if t == "Value" else 1.0 for t in titles]
    unit = (CONTENT_WIDTH - sum(label_w)) / sum(weights)
    tbl = Table(data, colWidths=label_w + [w * unit for w in weights], repeatRows=2, hAlign='LEFT')
    tbl.setStyle(TableStyle(style + [('TOPPADDING', (0, 2), (-1, -1), 1.5),
                                     ('BOTTOMPADDING', (0, 2), (-1, -1), 1.5)]))
    return tbl


def create_pdf(target, pdf_path, as_at, window, lmonths, rate_source, source_files):
    doc = SimpleDocTemplate(str(pdf_path), pagesize=PAGE_SIZE,
                            leftMargin=PAGE_MARGIN, rightMargin=PAGE_MARGIN,
                            topMargin=PAGE_MARGIN, bottomMargin=PAGE_MARGIN,
                            title=DPD_REPORT_TITLE)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('T', parent=styles['Title'], fontSize=16, spaceAfter=2,
                                 alignment=TA_CENTER, fontName='Helvetica-Bold',
                                 textColor=TITLE_NAVY)
    head_style = ParagraphStyle('H', parent=styles['Heading2'], fontSize=12, spaceAfter=2,
                                spaceBefore=2, alignment=TA_LEFT, fontName='Helvetica-Bold',
                                textColor=TABLE_TITLE_BLUE)
    note_style = ParagraphStyle('N', parent=styles['Normal'], fontSize=7, leading=8.5,
                                textColor=colors.grey, spaceAfter=3)
    rate_note = (f"Rate (Auto Finance) = {rate_source} ({BDO_PREFIX}) weighted by LEASE AMOUNT. "
                 if rate_source
                 else "Rate (Auto Finance): rate column not found, left blank. ")
    notes = (f"* Facilities activated (CONEX_DATE) {window_text(*window, sep=' to ')}. "
             f"No. = facilities; Value = sum of LEASE AMOUNT; {rate_note}Collection Ratio = sum "
             f"{COLLECTION_COL} / sum {DUE_COL} over the LENDING month-ends {lmonths[0].strftime('%b-%Y')} "
             f"to {lmonths[-1].strftime('%b-%Y')}. 90 DPD = DPD_Adj &ge; {DPD_THRESHOLD} "
             f"({as_at - pd.Timedelta(days=1):%B %Y}); Ratio = 90 DPD No. / No. Analysis of 90 DPD = number of facilities "
             f"by DPD_Adj (0-29, 30-59, 60-89, 90+). Decision: {CS_DECISION} (Credit Score), "
             f"else {RBP_DECISION} (RBP), else Manual.<br/>* " + customer_rule_text(as_at)
             + f" Source: {', '.join(Path(f).name for f in source_files)}.")
    elements = []
    for i, product in enumerate(PRODUCTS):
        if i:
            elements.append(PageBreak())
        pdata = target[target["PRODUCT_CATEGORY"] == product]
        elements += [Paragraph(DPD_REPORT_TITLE, title_style),
                     Paragraph(f"<u>{product} (as at {ordinal_date(as_at)})</u> &nbsp; "
                               f"{window_text(*window)} (Activation)", head_style),
                     Paragraph(notes, note_style),
                     _product_table(product, pdata)]
    doc.build(elements)
    log.info(f"PDF saved: {pdf_path}")
    return pdf_path

# ============================================================
# MAIN
# ============================================================

def parse_args():
    ap = argparse.ArgumentParser(description=DPD_REPORT_TITLE)
    ap.add_argument("--as-at-month", help="YYYY-MM (default: the month before the run month)")
    ap.add_argument("--lending-dir", default=str(LENDING_DIR),
                    help="folder with the monthly LENDING csvs (collection / due)")
    ap.add_argument("--activation-file", help=f"{BDO_PREFIX} for the activation window "
                                               f"(default: newest with that exact date range)")
    ap.add_argument("--dpd-file", help=f"DPD file (default: newest in {DPD_DIR})")
    ap.add_argument("--cs-file", help="Credit Score Decision Report (default: latest)")
    ap.add_argument("--rbp-file", help="RBP Security Approved (default: latest)")
    ap.add_argument("--history-file", help=f"{BDO_PREFIX} history (default: latest multi-month)")
    ap.add_argument("--no-move", action="store_true", help="leave downloads in data/ (testing)")
    return ap.parse_args()


def main():
    args = parse_args()
    setup_logging("rental_dpd_summary")
    t0 = time.time()
    try:
        # 1. Period: data month A, activation window, "as at" date (1st of the run month)
        lending_dir = Path(args.lending_dir)
        _, _, _, as_at_month, window = dpd_period(args.as_at_month)
        as_at = as_at_month.end_time.normalize()              # last day of A (data date)
        report_as_at = as_at + pd.Timedelta(days=1)          # shown as "as at 1st <run month>"
        log.info(f"{DPD_REPORT_TITLE} | as at {report_as_at:%d-%b-%Y} (data {as_at_month}) | "
                 f"activations {window[0]} to {window[1]}")

        # 2. Input files (newest of each, unless given on the command line)
        history = args.history_file or find_history_file(window)
        try:
            activations = args.activation_file or find_activation_file(window)
        except FileNotFoundError as e:
            log.warning(f"{e}. Taking the activations from the history file "
                        f"{Path(history).name} instead ({BDO_START_COL} in the window)")
            activations = history
        files = {"CREDIT_SCORE": args.cs_file or find_latest("CREDIT_SCORE"),
                 "RBP": args.rbp_file or find_latest("RBP"),
                 "ACTIVATIONS": activations,
                 "HISTORY": history,
                 "DPD": args.dpd_file or find_dpd_file()}
        for k, f in files.items():
            log.info(f"Input {k:<14}: {Path(f).name}")

        # 3-7. Activations, decisions, categories, LENDING, DPD
        df, target, lmonths, rate_source = prepare(files, lending_dir, as_at_month)

        # 8. Metrics (logged per product; the tables are built in step 9)
        for product in PRODUCTS + ["Total"]:
            d = target if product == "Total" else target[target["PRODUCT_CATEGORY"] == product]
            with_rate = product in RATE_PRODUCTS
            titles = [t for _, cols in column_groups(with_rate) for t in cols]
            log.info(f"Summary {product:<12} " + " | ".join(
                f"{t}: {fmt(k, v)}" for t, (k, v) in zip(titles, metrics(d, with_rate))))

        # 9. Outputs: preprocessed data, PDF, Excel; then move the downloads used
        out_dir = REPORTS_DIR / datetime.now().strftime("%Y-%m-%d")
        out_data = out_dir / "data"
        out_data.mkdir(parents=True, exist_ok=True)
        period = window_text(*window, sep=" to ")
        tag = f"{window[0].strftime('%B-%Y')}_to_{window[1].strftime('%B-%Y')}"
        drop = ["_KEY", "_MONTH"]
        df.drop(columns=drop).to_excel(out_data / f"Rental_DPD_Preprocessed_All_{tag}.xlsx", index=False)
        target.drop(columns=drop).to_excel(out_data / f"Rental_DPD_Preprocessed_AF_3W_2W_{tag}.xlsx",
                                           index=False)

        pdf_path = out_dir / f"{DPD_FILE_NAME}_{tag}.pdf"
        xlsx_path = out_dir / f"{DPD_FILE_NAME}_{tag}.xlsx"
        create_pdf(target, pdf_path, report_as_at, window, lmonths, rate_source, files.values())
        create_excel(target, xlsx_path, report_as_at, window)

        if not args.no_move:
            # Manually shared files (DPD, LENDING) stay where they are
            move_inputs({k: v for k, v in files.items()
                         if k != "DPD" and not (k == "ACTIVATIONS" and v == files["HISTORY"])},
                        out_data)
        log.info(f"===== {DPD_REPORT_TITLE} completed in {time.time() - t0:.1f}s =====")

        # Lines read by main_orchestrator.py
        print(f"RECORDS : {len(target)}")
        print(f"AMOUNT  : {target[BDO_AMOUNT_COL].sum():.2f}")
        print(f"PERIOD  : {period}")
        print(f"PDF     : {pdf_path}")
        print(f"EXCEL   : {xlsx_path}")
    except Exception as e:
        log.error(f"Failed to generate {DPD_REPORT_TITLE}: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
