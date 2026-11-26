"""
Approved Facility_(AF_3W_2W)_Analysis - summary script (workflow_v1.txt).

PERIOD
  Set by --mode (report_period.py): daily, weekly, cumulative, monthly, past-n-months,
  custom. Tables are "as at" the last day of the period.

HOW IT WORKS (main() follows these steps)
  1. Period      start / end dates of the mode
  2. Inputs      newest Approval Report HO + Branch, Credit Score, RBP, BDO history
  3. Approvals   HO + Branch combined (headers / footers skipped), APPR_SOURCE added,
                 only STAT_CON = ACTIVE kept
  4. Decisions   Credit Score (FACNO -> CR_GRADE) first, then RBP (FACILITY_NO ->
                 AI_DECISION), else Manual; DECISION_MATCHED_SRC records which
  5. Categories  Product Category (Auto Finance / 3W / 2W / Other), Product Segment,
                 Customer Type (Existing / New, from the BDO history)
  6. Outputs     PDF + Excel (one sheet per product) in reports/<YYYY-MM-DD>/;
                 Preprocessed_All / Preprocessed_AF_3W_2W and the downloads used in
                 reports/<YYYY-MM-DD>/data/

This file also holds the shared helpers (file lookup, decision matching, categorisation,
PDF styles) that the two rental report scripts import.

    python approved_facility_summary.py --mode cumulative
    python approved_facility_summary.py --mode past-n-months --n-months 3
    python approved_facility_summary.py --mode custom --from-date 2026-10-01 --to-date 2026-10-31
"""
import os
import sys
import re
import glob
import shutil
import argparse
import logging
import time
from pathlib import Path
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (CondPageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                Table, TableStyle)

from report_period import (MODES, add_period_args, apply_period_args, get_period,
                           mode_tag as period_mode_tag, mode_text as describe_mode,
                           period_label, period_tag)
from log_config import setup_logging

# Module logger; handlers (console + daily log file) are set up in main()
# so importing REPORT_NAME from here does not start a log session.
log = logging.getLogger("approved_facility_summary")

# ============================================================
# REPORT NAME  (Report Title, Email Subject and output file names)
# ============================================================
REPORT_NAME = "Approved Facility_(AF_3W_2W)_Analysis"

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / "credentials.env")   # EXISTING_LOOKBACK_MONTHS, HISTORY_MAX_TENOR_MONTHS
DATA_DIR = BASE_DIR / "data"             # downloaders save here (staging)
REPORTS_DIR = BASE_DIR / "reports"       # reports/<YYYY-MM-DD>/ and .../data/

# key -> (staging sub-folder, file-name prefix)
SOURCES = {
    "HO":           ("APPROVAL_REPORT_data_HO",     "APPROVAL_REPORT_HO"),
    "BRANCH":       ("APPROVAL_REPORT_data_BRANCH", "APPROVAL_REPORT_BRANCH"),
    "CREDIT_SCORE": ("CREDIT_SCORE_DECISION_data",  "CREDIT_SCORE_DECISION"),
    "RBP":          ("RBP_SECURITY_APPROVED_data",  "RBP_SECURITY_APPROVED"),
    "HISTORY":      ("BRANCH_DISB_OLD_data",        "BRANCH DISBURSEMENT OLD"),   # optional
}
APPR_SOURCE = {"HO": "HO", "BRANCH": "Branch"}

# ---- Column names, after normalisation (upper case, spaces -> '_') ----
KEY_COL = "NEW_CONNO"                   # facility no. in the Approval Report
STATUS_COL = "STAT_CON"                 # 'STAT CON'
ACTIVE_VALUE = "ACTIVE"
AMOUNT_COL = "LEASE_AMOUNT"             # 'LEASE AMOUNT'
RATE_COL = "INT_RATE"                   # 'INT RATE'
CONEX_COL = "CONEX_DATE"                # disbursement date
ACC_OPEN_COL = "ACC_OPENING_DATE"       # 'ACC OPENING DATE': used when CONEX_DATE is blank
ACC_OPEN_DAYFIRST = True                # text dates in that column are DD-MM-YYYY
# First column found is used as the customer identifier / closure date
CUSTOMER_ID_CANDIDATES = ["NIC_NO", "NIC", "CUS_ID", "CUSTOMER_ID", "CLIENT_NIC", "CUSTOMER_CODE"]
CLOSURE_CANDIDATES = ["CLOSURE_DATE", "CLOSED_DATE", "CLOSE_DATE"]

# ---- Customer history: BRANCH DISBURSEMENT OLD (Existing / New Customer) ----
HISTORY_KEY = "FACNO"                   # matched to NEW_CONNO for the approved facilities
HISTORY_START_CANDIDATES = ["CONEX_DATE", "DISB_DATE", "DISBURSED_DATE", "DISBURSEMENT_DATE",
                            "ACTIVATION_DATE", "ACTIVATED_DATE", "ACC_OPENING_DATE",
                            "START_DATE"]
HISTORY_DAYFIRST = True                 # text dates in that export are DD-MM-YYYY

# How facilities are linked to one customer:
#   "CLIENT_CODE": client code cut out of the facility number (NEW_CONNO / FACNO),
#                  Excel: =TEXT(VALUE(MID(A1,7,LEN(A1)-12)),"0")
#   "NIC":         NIC_NO looked up in the history report (NEW_CONNO = FACNO)
CUSTOMER_KEY_METHOD = "CLIENT_CODE"
CLIENT_CODE_SKIP_START = 6              # characters before the client code
CLIENT_CODE_SKIP_END = 6                # characters after it
DATE_DAYFIRST = False                   # text dates: False = MM-DD-YYYY (Oracle BI prompt format)

CS_KEY, CS_DECISION = "FACNO", "CR_GRADE"               # Credit Score Decision Report
RBP_KEY, RBP_DECISION = "FACILITY_NO", "AI_DECISION"     # RBP Security Approved

# ---- Decision mapping ----
DECISION_ORDER = ["Green", "Yellow", "Red", "Manual"]
UNMATCHED_DECISION = "Manual"
# Exact raw value (case-insensitive) -> decision. Checked before the keywords below.
# 'Light Red' / 'Dark Red' -> Red via the RED keyword below.
# Accepted / Approved / Rejected appear in the RBP AI_DECISION column; kept in both maps.
CS_DECISION_MAP = {
    "Accepted": "Green",
    "Approved": "Green",
    "Rejected": "Red",
}
RBP_DECISION_MAP = {
    "Accepted": "Green",
    "Approved": "Green",
    "Rejected": "Red",
}
# Whole-word keywords in the raw value (so 'REFERRED' does not match RED)
DECISION_KEYWORDS = [
    (re.compile(r"\bGREEN\b"), "Green"),
    (re.compile(r"\b(YELLOW|AMBER)\b"), "Yellow"),
    (re.compile(r"\bRED\b"), "Red"),
    (re.compile(r"\bMANUAL\b"), "Manual"),
]

SRC_CS = "Credit Score Decision Report"
SRC_RBP = "RBP Security Approved Report"
SRC_BOTH = "Both"
SRC_NONE = "None"
SRC_ORDER = [SRC_CS, SRC_RBP, SRC_BOTH, SRC_NONE]

# ---- Customer type ----
# Keep in step with branch_disb_data.py (both read these from credentials.env)
EXISTING_LOOKBACK_MONTHS = int(os.environ.get("EXISTING_LOOKBACK_MONTHS", "3"))
HISTORY_MAX_TENOR_MONTHS = int(os.environ.get("HISTORY_MAX_TENOR_MONTHS", "84"))
CUSTOMER_TYPES = ["Existing Customers", "New Customers"]
UNCLASSIFIED = "Customer Type N/A"      # used when the Approval Report has no customer ID

# ---- Products / segments ----
PRODUCTS = ["Auto Finance", "3W", "2W"]
SEGMENT_ORDER = ["HP/Lease", "Smart Draft", "Enhancement", "Pledge Loan", "Other"]
# Segment rows always shown (summary.xlsx); any other segment is added only when it has data
TEMPLATE_SEGMENTS = {
    ("Auto Finance", "Existing Customers"): ["HP/Lease", "Smart Draft", "Enhancement", "Pledge Loan"],
    ("3W", "Existing Customers"): ["HP/Lease", "Enhancement", "Pledge Loan"],
    ("2W", "Existing Customers"): ["HP/Lease", "Enhancement", "Pledge Loan"],
}
DEFAULT_TEMPLATE_SEGMENTS = ["HP/Lease"]              # New Customers
SEGMENT_DISPLAY = {"Enhancement": "Enhancements", "Pledge Loan": "Pledge loans"}
HP_LEASE_DISPLAY = {"Existing Customers": "New financing (Lease/HP)",
                    "New Customers": "Lease/HP"}

SMART_DRAFT_KEYS = [
    "int only", "intonly", "int-only", "interest only", "profit only",
    "intr only", "intrs only",
    "smart draft", "smart draf", "smart drft", "smrt draft", "smrt draf", "smrt drft",
]
# Whole words only: as substrings 'en' / 'enh' would hit ordinary words
ENHANCEMENT_RE = re.compile(
    r"\b(enhancement|enhancmnt|enhancemnt|enhncmnt|enhc|enhn|enh|en)\b", re.IGNORECASE)

# ------------------------------------------------------------
# STYLING (same as the Initiated Facility Details Report)
# ------------------------------------------------------------
PAGE_SIZE = landscape(letter)
PAGE_MARGIN = 0.4 * inch
CONTENT_WIDTH = PAGE_SIZE[0] - 2 * PAGE_MARGIN
FRAME_HEIGHT = PAGE_SIZE[1] - 2 * PAGE_MARGIN - 12

HEADER_BLUE = colors.HexColor('#4472C4')
SUBHDR_BLUE = colors.HexColor('#8EAADB')
TOTAL_FILL = colors.HexColor('#BDD7EE')
LABEL_FILL = colors.HexColor('#EAF0FA')
SUBTOTAL_FILL = colors.HexColor('#DEEAF6')
GRID_COLOR = colors.HexColor('#2F5496')
TITLE_NAVY = colors.HexColor('#1F3864')
TABLE_TITLE_BLUE = colors.HexColor('#2E4C7E')
DECISION_TINT = {"Green": '#E2EFDA', "Yellow": '#FFF2CC', "Red": '#FCE4D6', "Manual": '#EDEDED'}

# ============================================================
# HELPERS
# ============================================================

def _normalize(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).replace("_", " ").casefold().split())


def _key(series):
    """
    Clean key for facility / customer matching. Upper case, '.0' float suffix removed,
    then everything except letters and digits removed, so "'LE-12345", ' le 12345 ',
    'LE/12345' and 'LE12345.0' all become 'LE12345'.
    """
    key = series.astype(str).str.strip().str.upper().str.replace(r"\.0$", "", regex=True)
    key = key.str.replace(r"[^0-9A-Z]", "", regex=True)
    return key.where(~key.isin(["", "NAN", "NONE", "NAT", "NA"]))


def client_code(series):
    """
    Client code from a facility number, same as Excel
        =TEXT(VALUE(MID(A1, 7, LEN(A1) - 12)), "0")
    i.e. drop the first 6 and last 6 characters, read the middle as a number
    (leading zeros go). Blank when the number is too short or the middle is not numeric
    (Excel would show #VALUE!).
    A leading apostrophe (Excel text marker; present in the Approval, Branch Disbursement
    Old and Credit Score exports, not in RBP) is removed first, as Excel ignores it too.
    """
    raw = (series.astype(str).str.strip().str.strip("'\"`").str.strip()
           .str.replace(r"\.0$", "", regex=True))
    n = raw.str.len()
    mid = raw.str.slice(CLIENT_CODE_SKIP_START).where(
        n > CLIENT_CODE_SKIP_START + CLIENT_CODE_SKIP_END)
    mid = mid.str.slice(0, -CLIENT_CODE_SKIP_END).str.strip()
    numeric = mid.str.fullmatch(r"\d+", na=False)
    code = mid.where(numeric).str.lstrip("0")
    return code.where(code.isna() | (code != ""), "0")


def to_number(series):
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce").fillna(0).astype(float)


def to_date(series, dayfirst=None):
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    dayfirst = DATE_DAYFIRST if dayfirst is None else dayfirst
    return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst, format="mixed")


def first_present(df, candidates):
    return next((c for c in candidates if c in df.columns), None)


def weighted_rate(data):
    amount = data[AMOUNT_COL].sum()
    if not amount:
        return float('nan')
    return (data[AMOUNT_COL] * data[RATE_COL]).sum() / amount


def metrics(data):
    """(No. of files = count of NEW_CONNO, Lease Amount, WALR)."""
    return (int(data[KEY_COL].notna().sum()), float(data[AMOUNT_COL].sum()), weighted_rate(data))


def format_count(v):
    return "-" if not v else f"{int(v):,}"


def format_amount(v):
    return "-" if pd.isna(v) or v == 0 else f"{v:,.0f}"


def format_rate(v):
    return "-" if pd.isna(v) or v == 0 else f"{v:.2f}"


def ordinal_date(d):
    day = d.day
    suffix = "th" if 11 <= day % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix} {d:%B %Y}"

# ============================================================
# INPUT FILES
# ============================================================

BDO_RANGE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})")


def bdo_file_range(path):
    """(from, to) Timestamps from a BRANCH DISBURSEMENT OLD file name, or None."""
    m = BDO_RANGE_RE.search(Path(path).name)
    return tuple(pd.Timestamp(x) for x in m.groups()) if m else None


def is_month_bdo(path):
    """True for a single-calendar-month download (rental cohort file), not the history."""
    r = bdo_file_range(path)
    return bool(r) and r[0].day == 1 and r[1] == r[0] + pd.offsets.MonthEnd(0)


def is_history_bdo(path):
    """
    True for a customer-history download: a range of at least HISTORY_MAX_TENOR_MONTHS
    (one-month rental cohorts and the 12-month Rental Shortfall & 90 DPD activation
    download are not history). Files without a range in the name are accepted.
    """
    r = bdo_file_range(path)
    if not r:
        return True
    return r[1] >= r[0] + pd.DateOffset(months=HISTORY_MAX_TENOR_MONTHS) - pd.Timedelta(days=1)


def find_latest(source_key):
    """
    Newest <prefix>*.xlsx, searched in order:
      data/<subdir>/                 current downloaders
      <subdir>/                      old approval_report_data_HO/BRANCH.py layout
      Input_Data/<subdir>/           old CREDIT_SCORE / RBP script layout
      reports/*/data/                files already moved by an earlier run
    """
    subdir, prefix = SOURCES[source_key]
    folders = [DATA_DIR / subdir, BASE_DIR / subdir, BASE_DIR / "Input_Data" / subdir]
    patterns = [glob.escape(str(f)) + os.sep + f"{prefix}*.xlsx" for f in folders]
    patterns.append(glob.escape(str(REPORTS_DIR)) + os.sep + "*" + os.sep + "data"
                    + os.sep + f"{prefix}*.xlsx")
    for pattern in patterns:
        files = [Path(f) for f in glob.glob(pattern) if not os.path.basename(f).startswith("~$")]
        if source_key == "HISTORY":           # never a cohort / activation-window file
            files = [f for f in files if is_history_bdo(f)]
        if files:
            found = max(files, key=os.path.getmtime)
            if found.parent != DATA_DIR / subdir:
                log.info(f"{prefix}: using {found} (not in data/{subdir})")
            return found

    # Nothing found: say where we looked and what .xlsx files are there
    nearby = []
    for folder in folders:
        if folder.is_dir():
            nearby += [str(p) for p in folder.glob("*.xls*")]
    msg = (f"No {prefix}*.xlsx found. Looked in: "
           + "; ".join(str(f) for f in folders) + f"; {REPORTS_DIR}{os.sep}*{os.sep}data")
    if nearby:
        msg += f". Excel files present (name must start with {prefix}): {nearby[:10]}"
    raise FileNotFoundError(msg)


def read_report_excel(file_path, key_col):
    """Read an export: header = first row containing key_col; drop footer / Total rows."""
    raw = pd.read_excel(file_path, header=None)
    header_row = None
    for i, row in raw.iterrows():
        cells = row.astype(str).str.strip().str.upper().str.replace(r"\s+", "_", regex=True)
        if key_col in cells.values:
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"Header row with '{key_col}' not found in {file_path}")

    df = pd.read_excel(file_path, header=header_row)
    df.columns = (df.columns.astype(str).str.strip().str.upper()
                  .str.replace(r"\s+", "_", regex=True))
    df = df[_key(df[key_col]).notna()]
    df = df[~df[key_col].astype(str).str.upper().str.contains("TOTAL", na=False)]
    # Footer lines (e.g. 'Generated by ...') have text in the key column only
    df = df[df.drop(columns=[key_col]).notna().any(axis=1)].copy()
    return df


def require_cols(df, cols, label):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"{label}: column(s) {missing} not found. Columns: {list(df.columns)}")

# ============================================================
# CATEGORIZATION
# ============================================================

THREE_WHEELER_PRODUCTS = {"cash in hand", "ijarah small lease", "pledge loan",
                          "three wheeler-lease-brand new", "three wheeler-lease-registered"}
TWO_WHEELER_PRODUCTS = {"2w lease", "ijarah 2w lease"}
AUTO_FINANCE_PRODUCTS = {"diminishing musharaka", "hire purchase-registered",
                         "hire purchase-un-registered", "ijarah lease", "mini truck lease",
                         "murabaha", "other lease", "pledge loan", "vehicle loan-registered",
                         "vehicle loan-un-registered"}
AUTO_FINANCE_ASSETS = {"crew cab", "double cab", "jeep", "mini truck", "motor bus", "motor car",
                       "motor coach", "motor lorry", "motor truck", "motor van", "prime mover",
                       "single cab", "trailer"}


def categorize_product(product_name, asset_descr):
    """Product Category (proudct_categorization.txt section 1)."""
    product, asset = _normalize(product_name), _normalize(asset_descr)
    if product in THREE_WHEELER_PRODUCTS and asset == "three wheeler":
        return "3W"
    if product in TWO_WHEELER_PRODUCTS and asset == "motor cycle":
        return "2W"
    if product in AUTO_FINANCE_PRODUCTS and asset in AUTO_FINANCE_ASSETS:
        return "Auto Finance"
    return "Other"


def categorize_segment(product_name, scheme_description):
    """Product Segment (section 2). First rule that matches wins."""
    scheme, product = _normalize(scheme_description), _normalize(product_name)
    if any(k in scheme for k in SMART_DRAFT_KEYS):
        return "Smart Draft"
    if ENHANCEMENT_RE.search(scheme):
        return "Enhancement"
    if product == "pledge loan":
        return "Pledge Loan"
    if "hire purchase" in product or "lease" in product:
        return "HP/Lease"
    return "Other"

# ============================================================
# DATA PREPARATION
# ============================================================

def load_approvals(files):
    """files: {'HO': path, 'BRANCH': path}. Returns all rows (any status) with APPR_SOURCE."""
    frames = []
    for src, path in files.items():
        d = read_report_excel(path, KEY_COL)
        d["APPR_SOURCE"] = APPR_SOURCE[src]
        log.info(f"Approval Report {APPR_SOURCE[src]}: {len(d):,} rows from {Path(path).name}")
        frames.append(d)
    df = pd.concat(frames, ignore_index=True, sort=False)
    require_cols(df, [STATUS_COL, AMOUNT_COL, RATE_COL, "PRODUCT_NAME", "ASSET_DESCR"],
                 "Approval Report")
    for col in ["SCHEME_DESCRIPTION", CONEX_COL]:
        if col not in df.columns:
            log.warning(f"Approval Report: {col} not found; defaulting to blank.")
            df[col] = pd.NA
    df[AMOUNT_COL] = to_number(df[AMOUNT_COL])
    df[RATE_COL] = to_number(df[RATE_COL]).round(4)
    df[CONEX_COL] = to_date(df[CONEX_COL])

    # Blank CONEX_DATE -> ACC_OPENING_DATE
    if ACC_OPEN_COL in df.columns:
        df[ACC_OPEN_COL] = to_date(df[ACC_OPEN_COL], dayfirst=ACC_OPEN_DAYFIRST)
        blank = df[CONEX_COL].isna()
        df[CONEX_COL] = df[CONEX_COL].fillna(df[ACC_OPEN_COL])
        filled = int((blank & df[CONEX_COL].notna()).sum())
        log.info(f"Approval Report: {int(blank.sum()):,} blank {CONEX_COL}; {filled:,} filled "
                 f"from {ACC_OPEN_COL}, {int(df[CONEX_COL].isna().sum()):,} still blank")
    else:
        log.warning(f"Approval Report: {ACC_OPEN_COL} not found; blank {CONEX_COL} not filled.")

    df["_KEY"] = _key(df[KEY_COL])
    return df


def filter_active(df):
    status = df[STATUS_COL].astype(str).str.strip().str.upper()
    active = df[status == ACTIVE_VALUE].copy()
    log.info(f"{STATUS_COL} = {ACTIVE_VALUE}: {len(active):,} of {len(df):,} rows "
          f"(status counts {status.value_counts().to_dict()})")

    dup = active["_KEY"].duplicated(keep="first")
    if dup.any():
        log.warning(f"{int(dup.sum())} {KEY_COL}(s) appear more than once (HO and Branch?); "
              f"keeping the first: {active.loc[dup, KEY_COL].astype(str).tolist()[:20]}")
        active = active.loc[~dup].copy()
    return active


def map_decision(raw, mapping):
    if pd.isna(raw) or not str(raw).strip():
        return None
    text = " ".join(str(raw).upper().split())
    exact = {k.upper(): v for k, v in mapping.items()}
    if text in exact:
        return exact[text]
    for pattern, decision in DECISION_KEYWORDS:
        if pattern.search(text):
            return decision
    return None


def load_decisions(path, key_col, dec_col, mapping, label):
    """One row per facility: raw decision + mapped decision (None when unmapped)."""
    d = read_report_excel(path, key_col)
    require_cols(d, [key_col, dec_col], label)
    d = d[[key_col, dec_col]].copy()
    d["_KEY"] = _key(d[key_col])
    d["_DEC"] = d[dec_col].map(lambda v: map_decision(v, mapping))
    # Prefer a row with a usable decision; otherwise the last row
    d = (d.assign(_HAS=d["_DEC"].notna()).sort_values("_HAS", kind="stable")
          .drop_duplicates("_KEY", keep="last"))
    log.info(f"{label}: {len(d):,} facilities from {Path(path).name}; "
          f"{dec_col} values {d[dec_col].value_counts(dropna=False).to_dict()}")
    unmapped = d.loc[d["_DEC"].isna() & d[dec_col].notna(), dec_col].astype(str).unique()
    if len(unmapped):
        log.warning(f"{label}: {dec_col} value(s) not mapped to {DECISION_ORDER} "
              f"(treated as not matched): {sorted(unmapped)[:20]} "
              f"-> add them to {'CS' if label.startswith('Credit') else 'RBP'}_DECISION_MAP")
    return d.set_index("_KEY")[[dec_col, "_DEC"]]


# App Number exists only in the Approval Reports (not in Credit Score / RBP / Branch
# Disbursement Old), so it cannot be an alternative join key: no alternative is tried.
ALT_KEY_CANDIDATES = []


def log_match_rate(df, decisions, label, dec_key):
    """Log how many facilities find a row in a decision report; explain a poor match."""
    keys = df["_KEY"].dropna()
    hit = keys.isin(decisions.index)
    rate = hit.mean() if len(keys) else 0.0
    log.info(f"{label}: {int(hit.sum()):,} of {len(keys):,} facilities found "
             f"({rate:.0%}) on {KEY_COL} = {dec_key}")
    if rate >= 0.5:
        return
    log.warning(f"{label}: low match rate. Sample {KEY_COL}: {keys.head(5).tolist()} | "
                f"sample {dec_key}: {list(decisions.index[:5])}")
    for alt in ALT_KEY_CANDIDATES:
        if alt in df.columns:
            alt_keys = _key(df[alt]).dropna()
            alt_rate = alt_keys.isin(decisions.index).mean() if len(alt_keys) else 0.0
            log.warning(f"{label}: {alt} = {dec_key} would match {alt_rate:.0%}"
                        + (f" -> consider KEY_COL / join on {alt}" if alt_rate > rate else ""))


def match_decisions(df, cs, rbp):
    """Credit Score first, RBP when Credit Score has no usable decision."""
    df = df.copy()
    log_match_rate(df, cs, "Credit Score Decision Report", CS_KEY)
    log_match_rate(df, rbp, "RBP Security Approved", RBP_KEY)
    keys = df["_KEY"]
    df[f"CS_{CS_DECISION}"] = keys.map(cs[CS_DECISION])
    df["CS_DECISION"] = keys.map(cs["_DEC"])
    df[f"RBP_{RBP_DECISION}"] = keys.map(rbp[RBP_DECISION])
    df["RBP_DECISION"] = keys.map(rbp["_DEC"])

    has_cs, has_rbp = df["CS_DECISION"].notna(), df["RBP_DECISION"].notna()
    same = has_cs & has_rbp & (df["CS_DECISION"] == df["RBP_DECISION"])
    df["DECISION"] = (df["CS_DECISION"].where(has_cs, df["RBP_DECISION"])
                      .fillna(UNMATCHED_DECISION))
    df["DECISION_MATCHED_SRC"] = SRC_NONE
    df.loc[has_rbp, "DECISION_MATCHED_SRC"] = SRC_RBP
    df.loc[has_cs, "DECISION_MATCHED_SRC"] = SRC_CS
    df.loc[same, "DECISION_MATCHED_SRC"] = SRC_BOTH
    df["DECISION_CONFLICT"] = has_cs & has_rbp & ~same

    log.info(f"Decision source: {df['DECISION_MATCHED_SRC'].value_counts().to_dict()}")
    log.info(f"Decision: {df['DECISION'].value_counts().to_dict()}")
    if df["DECISION_CONFLICT"].any():
        log.warning(f"{int(df['DECISION_CONFLICT'].sum())} facility(ies) have different "
              f"decisions in the two reports; Credit Score decision used.")
    return df


def load_history(path):
    """
    BRANCH DISBURSEMENT OLD (customer history): one row per disbursed facility.
    Returns _HKEY (FACNO), _CUST (client code or NIC, see CUSTOMER_KEY_METHOD),
    _HSTART (disbursement date), _HCLOSE (closure date).
    """
    h = read_report_excel(path, HISTORY_KEY)
    cust_col = (HISTORY_KEY if CUSTOMER_KEY_METHOD == "CLIENT_CODE"
                else first_present(h, CUSTOMER_ID_CANDIDATES))
    start_col = first_present(h, HISTORY_START_CANDIDATES)
    close_col = first_present(h, CLOSURE_CANDIDATES)
    missing = [name for name, col in (("customer ID", cust_col),
                                      ("disbursement date", start_col)) if col is None]
    if missing:
        raise ValueError(f"History report {Path(path).name}: no {' / '.join(missing)} column "
                         f"(customer ID {CUSTOMER_ID_CANDIDATES}, disbursement date "
                         f"{HISTORY_START_CANDIDATES}). Columns: {list(h.columns)}")

    out = pd.DataFrame({
        "_HKEY": _key(h[HISTORY_KEY]),
        "_CUST": (client_code(h[HISTORY_KEY]) if CUSTOMER_KEY_METHOD == "CLIENT_CODE"
                  else _key(h[cust_col])),
        "_HSTART": to_date(h[start_col], HISTORY_DAYFIRST),
        "_HCLOSE": to_date(h[close_col], HISTORY_DAYFIRST) if close_col else pd.NaT,
    })
    bad_start = int(out["_HSTART"].isna().sum())
    out = out.dropna(subset=["_HKEY"]).drop_duplicates("_HKEY")
    cust_text = f"client code from {HISTORY_KEY}" if CUSTOMER_KEY_METHOD == "CLIENT_CODE" else cust_col
    log.info(f"Customer history: {len(out):,} facilities from {Path(path).name} | "
             f"{HISTORY_KEY} / {cust_text} / {start_col} / {close_col or 'no closure column'} | "
             f"disbursed {out['_HSTART'].min():%d-%b-%Y} to {out['_HSTART'].max():%d-%b-%Y} | "
             f"closed {int(out['_HCLOSE'].notna().sum()):,}")
    if bad_start:
        log.warning(f"Customer history: {bad_start} row(s) with an unreadable {start_col} "
                    f"(ignored for the Existing check)")
    if close_col is None:
        log.warning("Customer history: no closure date column; every earlier facility "
                    "counts as active.")
    return out


def check_history_coverage(path, conex):
    """Warn if the history file's From/To (from its file name) does not cover the windows."""
    m = re.search(r"(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})", Path(path).name)
    conex = conex.dropna()
    if not m or conex.empty:
        log.info("Customer history: date range not in file name; coverage not checked.")
        return
    h_from, h_to = (pd.Timestamp(x) for x in m.groups())
    need_from = conex.min() - pd.DateOffset(months=EXISTING_LOOKBACK_MONTHS)
    need_tenor = need_from - pd.DateOffset(months=HISTORY_MAX_TENOR_MONTHS)
    log.info(f"Customer history range {h_from:%d-%b-%Y} to {h_to:%d-%b-%Y} | needed: from "
             f"{need_tenor:%d-%b-%Y} (look-back + max tenor) to {conex.max():%d-%b-%Y}")
    if h_from > need_from:
        log.warning(f"Customer history starts {h_from:%d-%b-%Y}, after the look-back window "
                    f"start {need_from:%d-%b-%Y}: Existing Customers will be under-counted.")
    elif h_from > need_tenor:
        log.warning(f"Customer history starts {h_from:%d-%b-%Y}; long-tenor facilities "
                    f"disbursed before then (still active) are not seen.")
    if h_to < conex.max():
        log.warning(f"Customer history ends {h_to:%d-%b-%Y}, before the latest CONEX_DATE "
                    f"{conex.max():%d-%b-%Y}: NIC lookup will miss recent facilities.")


def classify_customers(active, approvals, history_file, as_at):
    """
    Existing Customer: on this facility's CONEX_DATE the customer held another CDB
    facility, or had one that closed within EXISTING_LOOKBACK_MONTHS before it,
    i.e. another facility disbursed before CONEX_DATE and not closed before
    CONEX_DATE - look-back. Otherwise New Customer.

    Customer key (CUSTOMER_KEY_METHOD):
      CLIENT_CODE  client code cut from NEW_CONNO (approved) and FACNO (history)
      NIC          NIC_NO from the history report, looked up via NEW_CONNO = FACNO
    Other facilities and their dates come from history_file (BRANCH DISBURSEMENT OLD),
    plus the ACTIVE Approval Report rows (start = CONEX_DATE, no closure date).
    Facilities without a customer key are shown as UNCLASSIFIED.
    """
    active = active.copy()
    by_code = CUSTOMER_KEY_METHOD == "CLIENT_CODE"
    frames = []
    if history_file:
        frames.append(load_history(history_file))
        check_history_coverage(history_file, active[CONEX_COL])
    else:
        log.warning("No BRANCH DISBURSEMENT OLD history: Existing is checked against this "
                    "period's approved facilities only, so it will be under-counted.")

    appr_cust = None if by_code else first_present(approvals, CUSTOMER_ID_CANDIDATES)
    if by_code or appr_cust:
        live = approvals[approvals[STATUS_COL].astype(str).str.strip().str.upper()
                         == ACTIVE_VALUE]
        frames.append(pd.DataFrame({
            "_HKEY": live["_KEY"],
            "_CUST": client_code(live[KEY_COL]) if by_code else _key(live[appr_cust]),
            "_HSTART": live[CONEX_COL], "_HCLOSE": pd.NaT}))

    if not frames:
        log.warning(f"No customer history and no customer ID column in the Approval Report; "
                    f"every facility is shown as '{UNCLASSIFIED}'.")
        active["CUSTOMER_TYPE"] = UNCLASSIFIED
        return active

    # History file first, so its closure dates win over the approval rows
    h = pd.concat(frames, ignore_index=True).drop_duplicates("_HKEY", keep="first")

    if by_code:
        cust = client_code(active[KEY_COL])
        sample = [f"{k} -> {c}" for k, c in zip(active[KEY_COL].astype(str), cust)][:3]
        log.info(f"Customer key: client code from {KEY_COL} / {HISTORY_KEY} "
                 f"(drop first {CLIENT_CODE_SKIP_START} and last {CLIENT_CODE_SKIP_END} "
                 f"characters), e.g. {sample}")
        source = f"client code from {KEY_COL}"
    else:
        cust = (_key(active[appr_cust]) if appr_cust
                else pd.Series(pd.NA, index=active.index, dtype=object))
        lookup = h.dropna(subset=["_CUST"]).set_index("_HKEY")["_CUST"]
        cust = cust.fillna(active["_KEY"].map(lookup))
        source = (f"Approval Report {appr_cust}" if appr_cust
                  else f"NIC via {KEY_COL} = {HISTORY_KEY}")

    found = int(cust.notna().sum())
    log.info(f"Customer key found for {found:,} of {len(active):,} approved facilities ({source})")
    if found < len(active):
        missing = active.loc[cust.isna().values, KEY_COL].astype(str).tolist()
        why = ("facility number too short or middle part not numeric" if by_code
               else "not found in the customer history")
        log.warning(f"{len(missing):,} approved facility(ies) without a customer key ({why}); "
                    f"shown as '{UNCLASSIFIED}'. Sample: {missing[:10]}")
    if by_code:
        bad_hist = int(h["_CUST"].isna().sum())
        if bad_hist:
            log.warning(f"{bad_hist:,} facility number(s) in the history / Approval Report gave "
                        f"no client code (ignored for the Existing check)")

    a = pd.DataFrame({"_KEY": active["_KEY"].values, "_CUST": cust.values,
                      "_CONEX": active[CONEX_COL].fillna(pd.Timestamp(as_at)).values})
    hh = h.dropna(subset=["_CUST", "_HSTART"])
    m = a.dropna(subset=["_CUST"]).merge(hh, on="_CUST")
    window_start = m["_CONEX"] - pd.DateOffset(months=EXISTING_LOOKBACK_MONTHS)
    m = m[(m["_HKEY"] != m["_KEY"]) & (m["_HSTART"] < m["_CONEX"])
          & (m["_HCLOSE"].isna() | (m["_HCLOSE"] >= window_start))]
    existing = set(m["_KEY"])

    active["CUSTOMER_TYPE"] = "New Customers"
    active.loc[active["_KEY"].isin(existing), "CUSTOMER_TYPE"] = "Existing Customers"
    active.loc[cust.isna().values, "CUSTOMER_TYPE"] = UNCLASSIFIED
    active["CLIENT_CODE" if by_code else "CUSTOMER_ID"] = cust.values
    log.info(f"Customer type: {active['CUSTOMER_TYPE'].value_counts().to_dict()}")
    return active


def prepare(files, as_at):
    """Steps 3-5: one row per active approved facility with every column the tables need."""
    # 3. Approvals: HO + Branch, ACTIVE only
    allrows = load_approvals({k: files[k] for k in ("HO", "BRANCH")})
    active = filter_active(allrows)

    # 4. Credit decision: Credit Score first, then RBP, else Manual
    cs = load_decisions(files["CREDIT_SCORE"], CS_KEY, CS_DECISION, CS_DECISION_MAP,
                        "Credit Score Decision Report")
    rbp = load_decisions(files["RBP"], RBP_KEY, RBP_DECISION, RBP_DECISION_MAP,
                         "RBP Security Approved")
    active = match_decisions(active, cs, rbp)

    # 5. Product category / segment and Existing / New customer
    active["PRODUCT_CATEGORY"] = [categorize_product(p, a) for p, a in
                                  zip(active["PRODUCT_NAME"], active["ASSET_DESCR"])]
    active["PRODUCT_SEGMENT"] = [categorize_segment(p, s) for p, s in
                                 zip(active["PRODUCT_NAME"], active["SCHEME_DESCRIPTION"])]
    log.info(f"Product Category: {active['PRODUCT_CATEGORY'].value_counts().to_dict()}")

    active = classify_customers(active, allrows, files.get("HISTORY"), as_at)

    target = active[active["PRODUCT_CATEGORY"].isin(PRODUCTS)].copy()
    log.info(f"Product Segment (AF/3W/2W): "
          f"{target.groupby(['PRODUCT_CATEGORY', 'PRODUCT_SEGMENT']).size().to_dict()}")
    return active.drop(columns=["_KEY"]), target.drop(columns=["_KEY"])

# ============================================================
# SUMMARY ROWS (shared by PDF and Excel)
# ============================================================

def customer_types(data):
    """Existing / New, plus the N/A bucket only when the data has it."""
    extra = [UNCLASSIFIED] if (data["CUSTOMER_TYPE"] == UNCLASSIFIED).any() else []
    return CUSTOMER_TYPES + extra


def segment_label(seg, ctype):
    if seg == "HP/Lease":
        return HP_LEASE_DISPLAY.get(ctype, "Lease/HP")
    return SEGMENT_DISPLAY.get(seg, seg)


def segments_for(product, ctype, data):
    base = TEMPLATE_SEGMENTS.get((product, ctype), DEFAULT_TEMPLATE_SEGMENTS)
    present = set(data["PRODUCT_SEGMENT"])
    return [s for s in SEGMENT_ORDER if s in base or s in present]


def summary_rows(product, pdata):
    """
    [(kind, label, decision, (files, amount, walr) | None)]
      type        customer-type heading
      decision    one decision row; label = segment on the first row of the segment
      subtotal    segment total
      type_total  customer-type total
      grand       product total
    """
    rows = []
    for ctype in customer_types(pdata):
        cdata = pdata[pdata["CUSTOMER_TYPE"] == ctype]
        rows.append(("type", ctype, "", None))
        for seg in segments_for(product, ctype, cdata):
            sdata = cdata[cdata["PRODUCT_SEGMENT"] == seg]
            label = segment_label(seg, ctype)
            for i, dec in enumerate(DECISION_ORDER):
                rows.append(("decision", label if i == 0 else "", dec,
                             metrics(sdata[sdata["DECISION"] == dec])))
            rows.append(("subtotal", f"Total - {label}", "", metrics(sdata)))
        rows.append(("type_total", f"Total - {ctype}", "", metrics(cdata)))
    rows.append(("grand", f"Total - {product}", "", metrics(pdata)))

    files = rows[-1][3][0]
    if files != len(pdata):
        log.warning(f"{product}: table total {files} != {len(pdata)} rows")
    return rows

# ============================================================
# EXCEL OUTPUT  (summary.xlsx layout: B..G, title rows 1-2, header row 3)
# ============================================================

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color.lstrip("#"))


def create_excel(target, xlsx_path, as_at, period_text):
    thin = Side(style="thin", color="2F5496")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill, hdr_font = _fill("#4472C4"), Font(bold=True, color="FFFFFF")
    fills = {"type": _fill("#EAF0FA"), "subtotal": _fill("#DEEAF6"),
             "type_total": _fill("#BDD7EE"), "grand": _fill("#BDD7EE")}
    count_fmt, amount_fmt, rate_fmt = '#,##0;-#,##0;"-"', '#,##0;-#,##0;"-"', '0.00;-0.00;"-"'

    wb = Workbook()
    wb.remove(wb.active)
    for product in PRODUCTS:
        ws = wb.create_sheet(product)
        pdata = target[target["PRODUCT_CATEGORY"] == product]
        ws["B1"] = f"{product} (as at {ordinal_date(as_at)})"
        ws["B1"].font = Font(bold=True, size=14, color="1F3864")
        ws["B2"] = period_text
        ws["B2"].font = Font(italic=True, color="2E4C7E")

        headers = ["Customer Type", "Product Segment", "Decision",
                   "No. of files", "Lease Amount (LKR)", "WALR"]
        for i, h in enumerate(headers):
            c = ws.cell(row=3, column=2 + i, value=h)
            c.fill, c.font, c.border = hdr_fill, hdr_font, border
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        r = 4
        seg_start = None
        for kind, label, dec, vals in summary_rows(product, pdata):
            cells = [ws.cell(row=r, column=col) for col in range(2, 8)]
            for c in cells:
                c.border = border
            if kind == "type":
                cells[0].value = label
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
            elif kind == "decision":
                if label:
                    cells[1].value = label
                    seg_start = r
                cells[2].value = dec
                cells[2].fill = _fill(DECISION_TINT[dec])
                if dec == DECISION_ORDER[-1]:
                    ws.merge_cells(start_row=seg_start, start_column=3, end_row=r, end_column=3)
                    ws.cell(row=seg_start, column=3).alignment = Alignment(vertical="center",
                                                                          wrap_text=True)
            elif kind == "subtotal":
                cells[1].value = label
                ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
            else:                                   # type_total / grand
                cells[0].value = label
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)

            if vals is not None:
                files, amount, walr = vals
                cells[3].value, cells[3].number_format = files, count_fmt
                cells[4].value, cells[4].number_format = round(amount, 2), amount_fmt
                cells[5].value = None if pd.isna(walr) else round(walr, 4)
                cells[5].number_format = rate_fmt

            if kind in fills:
                start = 1 if kind == "subtotal" else 0
                for c in cells[start:]:
                    c.fill = fills[kind]
                    c.font = Font(bold=True)
            r += 1

        for col, width in zip("ABCDEFG", [2, 22, 26, 11, 12, 20, 10]):
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "B4"

    wb.save(xlsx_path)
    log.info(f"Excel summary saved at: {xlsx_path}")
    return xlsx_path

# ============================================================
# PDF OUTPUT
# ============================================================

def _flow_height(flowable):
    _, h = flowable.wrap(CONTENT_WIDTH, FRAME_HEIGHT)
    return h + flowable.getSpaceBefore() + flowable.getSpaceAfter()


def add_block(elements, lead, table, min_body_rows=5):
    """Keep a table that fits on one page together with its heading."""
    lead_h = sum(_flow_height(f) for f in lead)
    _, table_h = table.wrap(CONTENT_WIDTH, FRAME_HEIGHT)
    if lead_h + table_h <= FRAME_HEIGHT:
        needed = lead_h + table_h
    else:
        needed = lead_h + sum(table._rowHeights[:1 + min_body_rows])
    elements.append(CondPageBreak(needed))
    elements.extend(lead)
    elements.append(table)


BASE_TABLE_STYLE = [
    ('GRID', (0, 0), (-1, -1), 0.5, GRID_COLOR),
    ('BOX', (0, 0), (-1, -1), 1.2, GRID_COLOR),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 8),
    ('LEADING', (0, 0), (-1, -1), 9.5),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('TOPPADDING', (0, 0), (-1, -1), 2),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
]


def _product_table(product, pdata):
    hdr = ["Customer Type", "Product Segment", "Decision",
           "No. of files", "Lease Amount (LKR)", "WALR (%)"]
    data = [hdr]
    style = list(BASE_TABLE_STYLE) + [
        ('BACKGROUND', (0, 0), (2, 0), HEADER_BLUE),
        ('BACKGROUND', (3, 0), (-1, 0), SUBHDR_BLUE),
        ('ALIGN', (3, 1), (4, -1), 'RIGHT'),
        ('ALIGN', (5, 1), (5, -1), 'CENTER'),
        ('ALIGN', (2, 1), (2, -1), 'CENTER'),
        ('RIGHTPADDING', (3, 1), (4, -1), 6),
    ]
    seg_start = None
    for kind, label, dec, vals in summary_rows(product, pdata):
        r = len(data)
        values = ["", "", ""] if vals is None else [format_count(vals[0]), format_amount(vals[1]),
                                                     format_rate(vals[2])]
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
        else:                                           # type_total / grand
            data.append([label, "", ""] + values)
            style += [('SPAN', (0, r), (2, r)), ('BACKGROUND', (0, r), (-1, r), TOTAL_FILL),
                      ('FONTNAME', (0, r), (-1, r), 'Helvetica-Bold')]
            if kind == "grand":
                style.append(('LINEABOVE', (0, r), (-1, r), 1.2, GRID_COLOR))

    widths = [1.5 * inch, 1.9 * inch, 0.8 * inch, 0.9 * inch, 1.5 * inch, 0.8 * inch]
    tbl = Table(data, colWidths=widths, repeatRows=1, hAlign='LEFT')
    tbl.setStyle(TableStyle(style))
    return tbl


def _grouped_table(row_labels, row_filters, target, label_header, value_fn, value_headers,
                   total_label, value_weights=None):
    """Rows x (Auto Finance | 3W | 2W | Total) with value_headers per group."""
    groups = PRODUCTS + ["Total"]
    nv = len(value_headers)
    hdr0 = [label_header] + sum([[g] + [""] * (nv - 1) for g in groups], [])
    hdr1 = [""] + value_headers * len(groups)
    data = [hdr0, hdr1]

    def group_data(g):
        return target if g == "Total" else target[target["PRODUCT_CATEGORY"] == g]

    for label, fn in list(zip(row_labels, row_filters)) + [(total_label, lambda d: d)]:
        row = [label]
        for g in groups:
            row += value_fn(fn(group_data(g)))
        data.append(row)

    total_row = len(data) - 1
    style = list(BASE_TABLE_STYLE) + [
        ('SPAN', (0, 0), (0, 1)),
        ('BACKGROUND', (0, 0), (-1, 1), HEADER_BLUE),
        ('BACKGROUND', (1, 1), (-1, 1), SUBHDR_BLUE),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
        ('FONTSIZE', (0, 1), (-1, 1), 7.5),
        ('ALIGN', (1, 2), (-1, -1), 'RIGHT'),
        ('RIGHTPADDING', (1, 2), (-1, -1), 5),
        ('BACKGROUND', (0, 2), (0, total_row - 1), LABEL_FILL),
        ('FONTNAME', (0, 2), (0, -1), 'Helvetica-Bold'),
        ('BACKGROUND', (0, total_row), (-1, total_row), TOTAL_FILL),
        ('FONTNAME', (0, total_row), (-1, total_row), 'Helvetica-Bold'),
    ]
    for g in range(len(groups)):
        c0 = 1 + g * nv
        style += [('SPAN', (c0, 0), (c0 + nv - 1, 0)),
                  ('LINEAFTER', (c0 + nv - 1, 0), (c0 + nv - 1, -1), 1.0, GRID_COLOR)]
    label_w = 1.8 * inch
    weights = value_weights or [1] * nv
    unit = (CONTENT_WIDTH - label_w) / (len(groups) * sum(weights))
    tbl = Table(data, colWidths=[label_w] + [w * unit for w in weights] * len(groups),
                repeatRows=2, hAlign='LEFT')
    tbl.setStyle(TableStyle(style))
    return tbl


def create_pdf(target, pdf_path, as_at, period_text, source_files):
    log.info(f"Generating PDF at: {pdf_path}")
    doc = SimpleDocTemplate(str(pdf_path), pagesize=PAGE_SIZE,
                            leftMargin=PAGE_MARGIN, rightMargin=PAGE_MARGIN,
                            topMargin=PAGE_MARGIN, bottomMargin=PAGE_MARGIN,
                            title=REPORT_NAME)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('MainTitle', parent=styles['Title'], fontSize=18, spaceAfter=4,
                                 alignment=TA_CENTER, fontName='Helvetica-Bold',
                                 textColor=TITLE_NAVY)
    table_title_style = ParagraphStyle('TableTitle', parent=styles['Heading2'], fontSize=12,
                                       spaceAfter=4, spaceBefore=6, alignment=TA_LEFT,
                                       fontName='Helvetica-Bold', textColor=TABLE_TITLE_BLUE)
    note_style = ParagraphStyle('Note', parent=styles['Normal'], fontSize=7.5, leading=9,
                                textColor=colors.grey, spaceAfter=3)
    source_style = ParagraphStyle('Source', parent=note_style, alignment=TA_CENTER, spaceAfter=8)

    def heading(text):
        return Paragraph(f"<u>{text}</u>", table_title_style)

    def note(lines):
        return Paragraph("<br/>".join(f"* {l}" for l in lines), note_style)

    elements = [
        Paragraph(REPORT_NAME, title_style),
        Paragraph(f"{period_text} | As at {ordinal_date(as_at)} | "
                  f"Generated: {datetime.now():%Y-%m-%d %H:%M}<br/>"
                  f"Source Files: {', '.join(Path(f).name for f in source_files)}", source_style),
    ]
    common_notes = [
        f"Approved facilities with {STATUS_COL.replace('_', ' ')} = {ACTIVE_VALUE} "
        f"(Approval Report HO + Branch).",
        "No. of files = count of NEW_CONNO; Lease Amount = sum of LEASE AMOUNT; "
        "WALR = Σ(Lease Amount × Int Rate) / Σ Lease Amount.",
    ]

    # 1. Overview
    ov_values = lambda d: [format_count(len(d)), format_amount(d[AMOUNT_COL].sum()),
                           format_rate(weighted_rate(d))]
    ctypes = customer_types(target)
    tbl = _grouped_table(ctypes,
                         [lambda d, c=c: d[d["CUSTOMER_TYPE"] == c] for c in ctypes],
                         target, "Customer Type", ov_values,
                         ["No. of files", "Lease Amount", "WALR (%)"], "Total",
                         value_weights=[1, 1.9, 0.9])
    add_block(elements, [heading("1. Overview"), note(common_notes + [
        f"Existing Customer: on the disbursement date (CONEX_DATE) the customer "
        f"({'client code from the facility no.' if CUSTOMER_KEY_METHOD == 'CLIENT_CODE' else 'NIC'}) held "
        f"another CDB facility, or had one that closed within the {EXISTING_LOOKBACK_MONTHS} "
        f"months before it (source: BRANCH DISBURSEMENT OLD). Otherwise New Customer."])], tbl)

    # 2-4. Product tables
    for i, product in enumerate(PRODUCTS, start=2):
        pdata = target[target["PRODUCT_CATEGORY"] == product]
        elements.append(Spacer(1, 8))
        add_block(elements, [heading(f"{i}. {product} (as at {ordinal_date(as_at)})"),
                             note([f"Decision: {CS_DECISION} from the Credit Score Decision "
                                   f"Report; if not matched, {RBP_DECISION} from RBP Security "
                                   f"Approved; not matched in either = {UNMATCHED_DECISION}."])],
                  _product_table(product, pdata))

    # 5. Decision match source
    elements.append(Spacer(1, 8))
    src_values = lambda d: [format_count(len(d)), format_amount(d[AMOUNT_COL].sum())]
    tbl = _grouped_table(SRC_ORDER,
                         [lambda d, s=s: d[d["DECISION_MATCHED_SRC"] == s] for s in SRC_ORDER],
                         target, "Decision Matched Source", src_values,
                         ["No. of files", "Lease Amount"], "Total")
    add_block(elements, [heading(f"{len(PRODUCTS) + 2}. Decision Match Source"),
                         note(["Both = same decision in both reports. "
                               "None = not matched in either report (shown as Manual)."])], tbl)

    doc.build(elements)
    log.info(f"PDF summary saved at: {pdf_path}")
    return pdf_path

# ============================================================
# MAIN
# ============================================================

def move_inputs(files, out_data_dir):
    """Move downloads used for this run from data/ (staging) into reports/<date>/data/."""
    moved = {}
    for key, path in files.items():
        path = Path(path)
        dest = out_data_dir / path.name
        try:
            in_staging = DATA_DIR.resolve() in path.resolve().parents
        except OSError:
            in_staging = False
        if in_staging and path.resolve() != dest.resolve():
            shutil.move(str(path), str(dest))
            moved[key] = dest
            log.info(f"Moved {path.name} -> {dest}")
        else:
            moved[key] = path
            log.info(f"{path.name} not in data/ staging; left in place")
    return moved


def parse_args():
    ap = argparse.ArgumentParser(description=REPORT_NAME)
    ap.add_argument("--mode", choices=MODES, default=os.environ.get("REPORT_MODE", "cumulative"))
    ap.add_argument("--n-months", type=int, default=int(os.environ.get("REPORT_N_MONTHS", "3")),
                    help="months for past-n-months")
    add_period_args(ap)                       # --from-date / --to-date for --mode custom
    ap.add_argument("--ho-file", help="Approval Report HO (default: latest)")
    ap.add_argument("--branch-file", help="Approval Report Branch (default: latest)")
    ap.add_argument("--cs-file", help="Credit Score Decision Report (default: latest)")
    ap.add_argument("--rbp-file", help="RBP Security Approved (default: latest)")
    ap.add_argument("--history-file",
                    help="BRANCH DISBURSEMENT OLD export for the Existing/New customer check "
                         "(default: latest in data/BRANCH_DISB_OLD_data)")
    ap.add_argument("--no-move", action="store_true",
                    help="leave the input files in data/ (testing)")
    args = ap.parse_args()
    if args.mode == "custom" and not (args.from_date or args.to_date):
        # Launched by the orchestrator without explicit dates: take them from the env
        args.from_date = os.environ.get("REPORT_FROM_DATE")
        args.to_date = os.environ.get("REPORT_TO_DATE")
    apply_period_args(ap, args)
    return args


def log_product_totals(target):
    """One line per product (and total) so each run's figures are in the log."""
    for product in PRODUCTS + ["Total"]:
        d = target if product == "Total" else target[target["PRODUCT_CATEGORY"] == product]
        files, amount, walr = metrics(d)
        by_type = d["CUSTOMER_TYPE"].value_counts().to_dict()
        log.info(f"Summary {product:<12}: files {files:>6,} | lease amount {amount:>18,.2f} | "
                 f"WALR {walr:6.2f} | {by_type}")


def main():
    args = parse_args()
    setup_logging("approved_facility_summary")
    t0 = time.time()
    try:
        # 1. Period
        start, end = get_period(args.mode, args.n_months)
        label = period_label(start, end)
        mode_text = describe_mode(args.mode, args.n_months)
        period_text = f"Period: {label} ({mode_text})"
        log.info(f"{REPORT_NAME} | mode {args.mode} | {start} to {end}")

        # 2. Input files (newest of each, unless given on the command line)
        files = {
            "HO": args.ho_file or find_latest("HO"),
            "BRANCH": args.branch_file or find_latest("BRANCH"),
            "CREDIT_SCORE": args.cs_file or find_latest("CREDIT_SCORE"),
            "RBP": args.rbp_file or find_latest("RBP"),
        }
        # Customer history (BRANCH DISBURSEMENT OLD) is optional
        if args.history_file:
            files["HISTORY"] = args.history_file
        else:
            try:
                files["HISTORY"] = find_latest("HISTORY")
            except FileNotFoundError:
                log.warning("No BRANCH DISBURSEMENT OLD file found; Existing/New Customer "
                            "split needs it (run branch_disb_data.py).")
        for k, f in files.items():
            if not os.path.exists(f):
                raise FileNotFoundError(f"{k} file not found: {f}")
            mtime = datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M")
            log.info(f"Input {k:<12}: {Path(f).name} (modified {mtime}, "
                     f"{os.path.getsize(f):,} bytes)")

        # 3-5. Approvals, decisions, categories
        active, target = prepare(files, end)
        if target.empty:
            log.warning("No Auto Finance / 3W / 2W facilities in the data; tables will be empty.")
        log_product_totals(target)

        # 6. Outputs: preprocessed data, PDF, Excel; then move the downloads used
        out_dir = REPORTS_DIR / datetime.now().strftime("%Y-%m-%d")
        out_data = out_dir / "data"
        out_data.mkdir(parents=True, exist_ok=True)
        mode_tag = period_mode_tag(args.mode, args.n_months)
        tag = f"{mode_tag}_{period_tag(start, end)}"

        all_path = out_data / f"Preprocessed_All_{tag}.xlsx"
        target_path = out_data / f"Preprocessed_AF_3W_2W_{tag}.xlsx"
        active.to_excel(all_path, index=False)
        target.to_excel(target_path, index=False)
        log.info(f"Preprocessed data saved to {out_data}: {all_path.name} ({len(active):,} rows), "
                 f"{target_path.name} ({len(target):,} rows)")

        pdf_path = out_dir / f"{REPORT_NAME}_{tag}.pdf"
        xlsx_path = out_dir / f"{REPORT_NAME}_{tag}.xlsx"
        create_pdf(target, pdf_path, end, period_text, files.values())
        create_excel(target, xlsx_path, end, period_text)

        if not args.no_move:
            move_inputs(files, out_data)
        else:
            log.info("--no-move: input files left in place")
        log.info(f"===== {REPORT_NAME} completed in {time.time() - t0:.1f}s =====")

        # Lines read by main_orchestrator.py
        print(f"RECORDS : {len(target)}")
        print(f"AMOUNT  : {target[AMOUNT_COL].sum():.2f}")
        print(f"PERIOD  : {label}")
        print(f"PDF     : {pdf_path}")
        print(f"EXCEL   : {xlsx_path}")
    except Exception as e:
        log.error(f"Failed to generate {REPORT_NAME}: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
