"""
Rental Shortfall Report (AF / 3W / 2W) - workflow_v2.txt.

AS-AT DATE
  Runs on the last day of each month. The report covers the run month and the 2 months
  before it, "as at" the last day of the latest LENDING month (= the run day).
  Run 31 Oct 2026 -> months Aug, Sep, Oct 2026, "as at 31st October 2026".
  (The Rental Shortfall & 90 DPD Analysis instead is "as at the 1st of the next month".)

HOW IT WORKS (main() follows these steps)
  1. LENDING     monthly portfolio csvs in data/LENDING_data (header = the row with MONTHEND
                 and NEW_CONNO); the latest 3 months found set the report months
  2. Inputs      Credit Score + RBP decisions, BDO history (Existing/New), and one
                 BRANCH DISBURSEMENT OLD per month (else that month from the history file)
  3. Cohorts     facilities disbursed in each month (= activation month)
  4. Decisions   Credit Score first, then RBP, else Manual
  5. Categories  Product (AF / 3W / 2W), segment, Existing / New (v2 rule: another facility
                 open on the run date, or taken in the 12 months before CONEX_DATE)
  6. LENDING     merge each cohort with THE SAME MONTH's portfolio (NEW_CONNO = FACNO)
  7. Metrics     per product / customer type / segment / decision and month:
                   No. = facilities, Value = sum LEASE_AMOUNT,
                   Collection Ratio = sum COLLECTION_FOR_THE_MONTH / sum DUE_FOR_THE_MONTH
  8. Outputs     PDF (one page per product) + Excel (one sheet per product) in reports/<date>/

    python rental_shortfall_summary.py
    python rental_shortfall_summary.py --lending-dir "D:/LENDING_data" --no-move
"""
import os
import re
import csv
import sys
import glob
import argparse
import logging
import time
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle

import approved_facility_summary as afs
from approved_facility_summary import (
    BASE_TABLE_STYLE, CS_DECISION, CS_DECISION_MAP, CS_KEY, CUSTOMER_TYPES, DATA_DIR,
    DECISION_ORDER, DECISION_TINT, GRID_COLOR, HEADER_BLUE, HISTORY_DAYFIRST, HISTORY_KEY,
    LABEL_FILL, PAGE_MARGIN, PAGE_SIZE, CONTENT_WIDTH, PRODUCTS, RBP_DECISION,
    RBP_DECISION_MAP, RBP_KEY, REPORTS_DIR, SEGMENT_ORDER, SUBHDR_BLUE, SUBTOTAL_FILL,
    TABLE_TITLE_BLUE, TITLE_NAVY, TOTAL_FILL, UNCLASSIFIED, _fill, _key, categorize_product,
    categorize_segment, client_code, customer_types, find_latest, first_present,
    format_amount, format_count, is_month_bdo, bdo_file_range, load_decisions,
    load_history, match_decisions, move_inputs, ordinal_date, read_report_excel,
    require_cols, segment_label, to_date, to_number, CLOSURE_CANDIDATES,
)
from log_config import setup_logging

log = logging.getLogger("rental_shortfall_summary")

# ============================================================
# REPORT NAME  (dedicated: Report Title, Email Subject, output file name)
# ============================================================
RENTAL_REPORT_TITLE = "Rental Shortfall Report"
RENTAL_EMAIL_SUBJECT = "Rental Shortfall Report"
RENTAL_FILE_NAME = "Rental Shortfall Report"

# ============================================================
# CONFIGURATION
# ============================================================
LENDING_DIR = DATA_DIR / "LENDING_data"
LENDING_MONTHS = 3                      # the latest 3 LENDING months found (run month + 2 before)
LENDING_KEY = "NEW_CONNO"               # facility number in the portfolio files
MONTHEND_COL = "MONTHEND"
COLLECTION_COL = "COLLECTION_FOR_THE_MONTH"
DUE_COL = "DUE_FOR_THE_MONTH"
GROSS_RENTAL_COL = "GROSS_RENTAL"
LENDING_COLS = [COLLECTION_COL, DUE_COL, GROSS_RENTAL_COL]
MONTHEND_DAYFIRST = True                # text MONTHEND like 31-10-2026

BDO_KEY = HISTORY_KEY                   # FACNO in BRANCH DISBURSEMENT OLD
BDO_SUBDIR, BDO_PREFIX = afs.SOURCES["HISTORY"]
BDO_START_COL = "CONEX_DATE"            # disbursement date in BRANCH DISBURSEMENT OLD
BDO_AMOUNT_COL = "LEASE_AMOUNT"         # 'Value' = sum of LEASE AMOUNT

# Customer type (customer_categorization_v2.txt)
RENTAL_LOOKBACK_MONTHS = 12             # "had a facility last 12 months before taking new"

# Segment rows always shown; other segments added only when they have data
RENTAL_TEMPLATE_SEGMENTS = {
    "Existing Customers": ["HP/Lease", "Smart Draft", "Enhancement", "Pledge Loan"],
}
RENTAL_DEFAULT_SEGMENTS = ["HP/Lease"]   # New Customers (and N/A)


def customer_rule_text(as_at):
    """v2 Existing/New rule, worded as tests A and B (PDF note and Excel sheets)."""
    return (f"Existing Customer if another facility of the same client code passes A or B "
            f"(else New). A: open on the as-at date ({ordinal_date(as_at)}). "
            f"B: disbursed in the {RENTAL_LOOKBACK_MONTHS} months before this facility's "
            f"CONEX_DATE.")

# ============================================================
# LENDING PORTFOLIOS
# ============================================================

def _read_csv_raw(path):
    """All rows as text; title / footer lines with fewer fields are padded, not an error."""
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, newline="", encoding=enc) as fh:
                rows = list(csv.reader(fh))
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f"{Path(path).name}: cannot read as UTF-8 or Windows-1252")
    width = max((len(r) for r in rows), default=0)
    raw = pd.DataFrame([r + [""] * (width - len(r)) for r in rows], dtype=object)
    return raw.replace(r"^\s*$", pd.NA, regex=True)


def read_lending_csv(path):
    """One portfolio csv: header = first row containing MONTHEND and NEW_CONNO."""
    raw = _read_csv_raw(path)
    header_row = None
    for i, row in raw.head(50).iterrows():
        cells = set(row.dropna().astype(str).str.strip().str.upper()
                    .str.replace(r"\s+", "_", regex=True))
        if MONTHEND_COL in cells and LENDING_KEY in cells:
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"{Path(path).name}: no header row with {MONTHEND_COL} and {LENDING_KEY}")
    cols = (raw.iloc[header_row].astype(str).str.strip().str.upper()
            .str.replace(r"\s+", "_", regex=True))
    df = raw.iloc[header_row + 1:].copy()
    df.columns = cols
    require_cols(df, [LENDING_KEY, MONTHEND_COL] + LENDING_COLS, f"LENDING {Path(path).name}")
    df = df[_key(df[LENDING_KEY]).notna()].copy()
    df["_KEY"] = _key(df[LENDING_KEY])
    df["_MONTHEND"] = to_date(df[MONTHEND_COL], dayfirst=MONTHEND_DAYFIRST)
    for c in LENDING_COLS:
        df[c] = to_number(df[c])
    df["_SOURCE_FILE"] = Path(path).name
    return df


def load_lending(lending_dir=LENDING_DIR, months=LENDING_MONTHS):
    """
    All portfolio csvs -> one frame with _MONTH (Period 'M').
    Keeps the latest `months` months found; one row per facility per month
    (duplicates summed).
    """
    files = sorted(Path(f) for f in glob.glob(glob.escape(str(lending_dir)) + os.sep + "*.csv"))
    if not files:
        raise FileNotFoundError(f"No .csv files in {lending_dir}")
    frames = []
    for f in files:
        d = read_lending_csv(f)
        found = sorted(d["_MONTHEND"].dropna().dt.to_period("M").unique())
        log.info(f"LENDING {f.name}: {len(d):,} rows | MONTHEND month(s) "
                 f"{[str(m) for m in found] or 'unreadable'}")
        bad = int(d["_MONTHEND"].isna().sum())
        if bad:
            log.warning(f"LENDING {f.name}: {bad:,} row(s) with an unreadable {MONTHEND_COL} (ignored)")
        frames.append(d.dropna(subset=["_MONTHEND"]))
    df = pd.concat(frames, ignore_index=True)
    df["_MONTH"] = df["_MONTHEND"].dt.to_period("M")

    all_months = sorted(df["_MONTH"].unique())
    keep = all_months[-months:]
    if len(all_months) > months:
        log.warning(f"LENDING: {len(all_months)} months found {[str(m) for m in all_months]}; "
                    f"using the latest {months}: {[str(m) for m in keep]}")
    if len(keep) < months:
        log.warning(f"LENDING: only {len(keep)} month(s) found {[str(m) for m in keep]} "
                    f"(expected {months})")
    gaps = [str(b) for a, b in zip(keep, keep[1:]) if (b - a).n != 1]
    if gaps:
        log.warning(f"LENDING: months are not consecutive (gap before {gaps})")
    df = df[df["_MONTH"].isin(keep)]

    dup = df.duplicated(["_MONTH", "_KEY"], keep=False)
    if dup.any():
        log.warning(f"LENDING: {int(dup.sum()):,} rows share a {LENDING_KEY} within a month; "
                    f"their {', '.join(LENDING_COLS)} are summed")
    agg = (df.groupby(["_MONTH", "_KEY"], as_index=False)[LENDING_COLS].sum())
    return agg, keep, df["_MONTHEND"].max()


def lending_period(lending_dir=LENDING_DIR, months=LENDING_MONTHS):
    """(first day, last day, [months]) of the portfolio months, for the orchestrator."""
    _, keep, _ = load_lending(lending_dir, months)
    first, last = keep[0], keep[-1]
    return first.start_time.date(), last.end_time.date(), keep

# ============================================================
# BRANCH DISBURSEMENT OLD (cohorts + history)
# ============================================================

def _bdo_files():
    folders = [DATA_DIR / BDO_SUBDIR, afs.BASE_DIR / BDO_SUBDIR]
    files = []
    for folder in folders:
        files += [Path(f) for f in glob.glob(glob.escape(str(folder)) + os.sep + f"{BDO_PREFIX}*.xlsx")]
    files += [Path(f) for f in glob.glob(glob.escape(str(REPORTS_DIR)) + os.sep + "*" + os.sep
                                         + "data" + os.sep + f"{BDO_PREFIX}*.xlsx")]
    return [f for f in files if not f.name.startswith("~$")]


def find_cohort_files(months, history_file=None):
    """
    Newest one-month BRANCH DISBURSEMENT OLD file for each month. A month without one
    falls back to history_file (facilities with CONEX_DATE in that month), if given.
    """
    by_month = {}
    for f in _bdo_files():
        if is_month_bdo(f):
            m = bdo_file_range(f)[0].to_period("M")
            if m in months and (m not in by_month or f.stat().st_mtime > by_month[m].stat().st_mtime):
                by_month[m] = f
    missing = [str(m) for m in months if m not in by_month]
    if missing and history_file:
        log.warning(f"No one-month {BDO_PREFIX} file for {missing}; taking those months' "
                    f"activations from {Path(history_file).name} ({BDO_START_COL} in the month)")
        for m in months:
            by_month.setdefault(m, Path(history_file))
        return by_month
    if missing:
        raise FileNotFoundError(
            f"No one-month {BDO_PREFIX} file for {missing} (name must contain "
            f"<YYYY-MM-01>_to_<YYYY-MM-last day>; download with branch_disb_data.py "
            f"--from-date / --to-date)")
    return by_month


def load_cohorts(cohort_files):
    """Facilities disbursed in each month; _MONTH = activation (cohort) month."""
    frames, cache = [], {}
    for month, path in sorted(cohort_files.items()):
        if path not in cache:
            cache[path] = read_report_excel(path, BDO_KEY)
        d = cache[path].copy()
        require_cols(d, [BDO_AMOUNT_COL, "PRODUCT_NAME", "ASSET_DESCR"], f"{BDO_PREFIX} {path.name}")
        if "SCHEME_DESCRIPTION" not in d.columns:
            log.warning(f"{path.name}: SCHEME_DESCRIPTION not found; segments use PRODUCT_NAME only")
            d["SCHEME_DESCRIPTION"] = pd.NA
        d[BDO_AMOUNT_COL] = to_number(d[BDO_AMOUNT_COL])
        if BDO_START_COL in d.columns:
            d[BDO_START_COL] = to_date(d[BDO_START_COL], HISTORY_DAYFIRST)
        else:
            log.warning(f"{path.name}: {BDO_START_COL} not found; using the 1st of {month}")
            d[BDO_START_COL] = pd.NaT
        outside = d[BDO_START_COL].notna() & (d[BDO_START_COL].dt.to_period("M") != month)
        if not is_month_bdo(path):              # history fallback: keep this month only
            d = d[~outside & d[BDO_START_COL].notna()].copy()
            outside = outside.loc[d.index]
        if outside.any():
            log.warning(f"{path.name}: {int(outside.sum()):,} facility(ies) with {BDO_START_COL} "
                        f"outside {month}; kept in the {month} cohort")
        d[BDO_START_COL] = d[BDO_START_COL].fillna(month.start_time)
        d["_MONTH"] = month
        d["_SOURCE_FILE"] = path.name
        log.info(f"Cohort {month}: {len(d):,} facilities from {path.name}")
        frames.append(d)
    df = pd.concat(frames, ignore_index=True, sort=False)
    df["_KEY"] = _key(df[BDO_KEY])
    dup = df["_KEY"].duplicated(keep="first")
    if dup.any():
        log.warning(f"{int(dup.sum()):,} {BDO_KEY}(s) in more than one cohort file; keeping the first")
        df = df[~dup].copy()
    return df


def cohort_history_frame(cohorts):
    """Cohort rows in the load_history() shape, so they count as 'other facilities' too."""
    close_col = first_present(cohorts, CLOSURE_CANDIDATES)
    return pd.DataFrame({
        "_HKEY": cohorts["_KEY"].values,
        "_CUST": client_code(cohorts[BDO_KEY]).values,
        "_HSTART": cohorts[BDO_START_COL].values,
        "_HCLOSE": (to_date(cohorts[close_col], HISTORY_DAYFIRST).values if close_col
                    else pd.NaT),
    })

# ============================================================
# CUSTOMER TYPE (customer_categorization_v2.txt)
# ============================================================

def classify_customers_v2(df, history_file, as_at):
    """
    Existing if another facility of the same client code
      (A) is open on the run date (as_at):
          disbursed on/before as_at and not closed before it, OR
      (B) was taken in the RENTAL_LOOKBACK_MONTHS before this facility's CONEX_DATE:
          CONEX_DATE - look-back <= its disbursement date < CONEX_DATE (open or closed).
    Else New. No client code -> UNCLASSIFIED.
    Test A needs every facility that can still be open on as_at, however old, hence
    the long BDO history (HISTORY_MAX_TENOR_MONTHS).
    """
    df = df.copy()
    as_at = pd.Timestamp(as_at)
    frames = []
    if history_file:
        frames.append(load_history(history_file))
        r = bdo_file_range(history_file)
        need_from = df[BDO_START_COL].min() - pd.DateOffset(months=RENTAL_LOOKBACK_MONTHS)
        if r:
            log.info(f"Customer history range {r[0]:%d-%b-%Y} to {r[1]:%d-%b-%Y} | needed: "
                     f"from {need_from:%d-%b-%Y} (12-month look-back) to {as_at:%d-%b-%Y} (run date)")
            if r[0] > need_from:
                log.warning("Customer history starts after the 12-month look-back: "
                            "Existing Customers will be under-counted.")
            if r[1] < as_at:
                log.warning("Customer history ends before the run date: facilities opened "
                            "after it are not seen for 'open on the run date'.")
    else:
        log.warning(f"No {BDO_PREFIX} history file: Existing is checked against these 3 "
                    f"cohorts only, so it will be under-counted.")
    frames.append(cohort_history_frame(df))
    h = (pd.concat(frames, ignore_index=True).drop_duplicates("_HKEY", keep="first")
         .dropna(subset=["_CUST", "_HSTART"]))

    cust = client_code(df[BDO_KEY])
    a = pd.DataFrame({"_KEY": df["_KEY"].values, "_CUST": cust.values,
                      "_CONEX": df[BDO_START_COL].values})
    m = a.dropna(subset=["_CUST"]).merge(h, on="_CUST")
    m = m[m["_HKEY"] != m["_KEY"]]
    active_as_at = (m["_HSTART"] <= as_at) & (m["_HCLOSE"].isna() | (m["_HCLOSE"] > as_at))
    window_start = m["_CONEX"] - pd.DateOffset(months=RENTAL_LOOKBACK_MONTHS)
    in_lookback = (m["_HSTART"] < m["_CONEX"]) & (m["_HSTART"] >= window_start)
    existing = set(m.loc[active_as_at | in_lookback, "_KEY"])
    keys_a = set(m.loc[active_as_at, "_KEY"])
    keys_b = set(m.loc[in_lookback, "_KEY"])
    log.info(f"Existing via A 'open on {as_at:%d-%b-%Y}': "
             f"{m.loc[active_as_at, '_KEY'].nunique():,} | via B 'taken in the "
             f"{RENTAL_LOOKBACK_MONTHS} months before CONEX': "
             f"{m.loc[in_lookback, '_KEY'].nunique():,} (a facility can be both)")

    df["CLIENT_CODE"] = cust.values
    # Which test made the facility Existing: "A+B", "A", "B" ("" for New / N/A)
    df["EXISTING_TEST"] = [("A+B" if k in keys_a and k in keys_b else "A" if k in keys_a
                            else "B" if k in keys_b else "") for k in df["_KEY"]]
    df["CUSTOMER_TYPE"] = "New Customers"
    df.loc[df["_KEY"].isin(existing), "CUSTOMER_TYPE"] = "Existing Customers"
    df.loc[cust.isna().values, "CUSTOMER_TYPE"] = UNCLASSIFIED
    no_code = int(cust.isna().sum())
    if no_code:
        log.warning(f"{no_code:,} facility(ies) without a client code (facility number too short "
                    f"or not numeric); shown as '{UNCLASSIFIED}'. Sample: "
                    f"{df.loc[cust.isna().values, BDO_KEY].astype(str).tolist()[:10]}")
    log.info(f"Customer type: {df['CUSTOMER_TYPE'].value_counts().to_dict()}")
    return df

# ============================================================
# PREPARATION
# ============================================================

def prepare(lending, months, as_at, cohort_files, files):
    """Steps 3-6: one row per activated facility with every column the tables need."""
    # 3. Cohorts: facilities disbursed in each month
    df = load_cohorts(cohort_files)

    # 4. Credit decision: Credit Score first, then RBP, else Manual
    cs = load_decisions(files["CREDIT_SCORE"], CS_KEY, CS_DECISION, CS_DECISION_MAP,
                        "Credit Score Decision Report")
    rbp = load_decisions(files["RBP"], RBP_KEY, RBP_DECISION, RBP_DECISION_MAP,
                         "RBP Security Approved")
    df = match_decisions(df, cs, rbp)          # also logs the match rate per report

    # 5. Product category / segment and Existing / New customer
    df["PRODUCT_CATEGORY"] = [categorize_product(p, a) for p, a in
                              zip(df["PRODUCT_NAME"], df["ASSET_DESCR"])]
    df["PRODUCT_SEGMENT"] = [categorize_segment(p, s) for p, s in
                             zip(df["PRODUCT_NAME"], df["SCHEME_DESCRIPTION"])]
    log.info(f"Product Category: {df['PRODUCT_CATEGORY'].value_counts().to_dict()}")

    df = classify_customers_v2(df, files.get("HISTORY"), as_at)

    # 6. Same month's portfolio: cohort month M <-> LENDING MONTHEND month M
    df = df.merge(lending, how="left", on=["_MONTH", "_KEY"], suffixes=("", "_LENDING"))
    found = df[DUE_COL].notna()
    for month in months:
        sel = df["_MONTH"] == month
        log.info(f"Cohort {month}: {int((sel & found).sum()):,} of {int(sel.sum()):,} facilities "
                 f"found in the {month} portfolio")
    if (~found).any():
        log.warning(f"{int((~found).sum()):,} facility(ies) not in their month's portfolio "
                    f"(counted in No./Value, not in the Collection Ratio). Sample: "
                    f"{df.loc[~found, BDO_KEY].astype(str).tolist()[:10]}")
    df["FOUND_IN_PORTFOLIO"] = found
    df["ACTIVATION_MONTH"] = df["_MONTH"].astype(str)

    target = df[df["PRODUCT_CATEGORY"].isin(PRODUCTS)].copy()
    log.info(f"Product Segment (AF/3W/2W): "
             f"{target.groupby(['PRODUCT_CATEGORY', 'PRODUCT_SEGMENT']).size().to_dict()}")
    return df, target

# ============================================================
# SUMMARY ROWS
# ============================================================

def month_metrics(data, months):
    """[(No., Value, Collection Ratio)] per month."""
    out = []
    for m in months:
        d = data[data["_MONTH"] == m]
        due = d[DUE_COL].sum()
        ratio = d[COLLECTION_COL].sum() / due if due else float("nan")
        out.append((len(d), float(d[BDO_AMOUNT_COL].sum()), ratio))
    return out


def rental_segments(ctype, data):
    base = RENTAL_TEMPLATE_SEGMENTS.get(ctype, RENTAL_DEFAULT_SEGMENTS)
    present = set(data["PRODUCT_SEGMENT"])
    return [s for s in SEGMENT_ORDER if s in base or s in present]


def rental_rows(product, pdata, months):
    """Same row kinds as the Approved Facility tables; values = per-month triples."""
    rows = []
    for ctype in customer_types(pdata):
        cdata = pdata[pdata["CUSTOMER_TYPE"] == ctype]
        rows.append(("type", ctype, "", None))
        for seg in rental_segments(ctype, cdata):
            sdata = cdata[cdata["PRODUCT_SEGMENT"] == seg]
            label = segment_label(seg, ctype)
            for i, dec in enumerate(DECISION_ORDER):
                rows.append(("decision", label if i == 0 else "", dec,
                             month_metrics(sdata[sdata["DECISION"] == dec], months)))
            rows.append(("subtotal", f"Total - {label}", "", month_metrics(sdata, months)))
        rows.append(("type_total", f"Total - {ctype}", "", month_metrics(cdata, months)))
    rows.append(("grand", "Total", "", month_metrics(pdata, months)))
    return rows


def format_ratio(v):
    return "-" if pd.isna(v) else f"{v:.1%}"


def month_title(m):
    return f"{m.strftime('%B %Y')} activations"


def subtitle(months):
    first, last = months[0], months[-1]
    span = (f"{first.strftime('%B')} to {last.strftime('%B %Y')}" if first.year == last.year
            else f"{first.strftime('%B %Y')} to {last.strftime('%B %Y')}")
    return f"First 3 rental posted ({span})"

# ============================================================
# EXCEL (Summary.xlsx layout)
# ============================================================

def create_excel(target, xlsx_path, months, as_at):
    thin = Side(style="thin", color="2F5496")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill, sub_fill = _fill("#4472C4"), _fill("#8EAADB")
    white = Font(bold=True, color="FFFFFF")
    fills = {"type": _fill("#EAF0FA"), "subtotal": _fill("#DEEAF6"),
             "type_total": _fill("#BDD7EE"), "grand": _fill("#BDD7EE")}
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    n_cols = 3 + 3 * len(months)
    last_col = 1 + n_cols                    # B = 2

    wb = Workbook()
    wb.remove(wb.active)
    for product in PRODUCTS:
        ws = wb.create_sheet(product)
        pdata = target[target["PRODUCT_CATEGORY"] == product]
        ws.cell(row=1, column=2, value=f"{product} (as at {ordinal_date(as_at)})").font = \
            Font(bold=True, size=14, color="1F3864")
        ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last_col)
        ws.cell(row=2, column=2, value=subtitle(months)).font = Font(italic=True, color="2E4C7E")
        ws.cell(row=3, column=2, value=customer_rule_text(as_at)).font = Font(italic=True, size=9,
                                                                             color="808080")
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col)

        # Header rows 4-6
        for col, text in zip((2, 3, 4), ("Customer Type", "Product Segment", "Decision")):
            ws.cell(row=4, column=col, value=text)
            ws.merge_cells(start_row=4, start_column=col, end_row=6, end_column=col)
        for i, m in enumerate(months):
            c0 = 5 + 3 * i
            ws.cell(row=4, column=c0, value=month_title(m))
            ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 2)
            ws.cell(row=5, column=c0, value="Rental shortfall")
            ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 1)
            ws.cell(row=5, column=c0 + 2, value="Collection Ratio")
            ws.merge_cells(start_row=5, start_column=c0 + 2, end_row=6, end_column=c0 + 2)
            ws.cell(row=6, column=c0, value="No.")
            ws.cell(row=6, column=c0 + 1, value="Value")
        for r in (4, 5, 6):
            for col in range(2, last_col + 1):
                c = ws.cell(row=r, column=col)
                c.fill = hdr_fill if (r == 4 or col < 5) else sub_fill
                c.font, c.alignment, c.border = white, center, border

        r = 7
        seg_start = None
        for kind, label, dec, vals in rental_rows(product, pdata, months):
            cells = [ws.cell(row=r, column=col) for col in range(2, last_col + 1)]
            for c in cells:
                c.border = border
            if kind == "type":
                cells[0].value = label
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
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
            else:
                cells[0].value = label
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)

            if vals is not None:
                for i, (no, value, ratio) in enumerate(vals):
                    c0 = 3 + 3 * i            # index into cells (B = 0)
                    cells[c0].value, cells[c0].number_format = no, '#,##0;-#,##0;"-"'
                    cells[c0 + 1].value = round(value, 2)
                    cells[c0 + 1].number_format = '#,##0;-#,##0;"-"'
                    cells[c0 + 2].value = None if pd.isna(ratio) else round(ratio, 4)
                    cells[c0 + 2].number_format = '0.0%'
            if kind in fills:
                start = 1 if kind == "subtotal" else 0
                for c in cells[start:]:
                    c.fill = fills[kind]
                    c.font = Font(bold=True)
            r += 1

        ws.column_dimensions["A"].width = 2
        for col, width in zip("BCD", (22, 26, 11)):
            ws.column_dimensions[col].width = width
        for i in range(len(months)):
            for j, width in enumerate((9, 18, 12)):
                ws.column_dimensions[get_column_letter(5 + 3 * i + j)].width = width
        ws.freeze_panes = "E7"

    wb.save(xlsx_path)
    log.info(f"Excel saved: {xlsx_path}")
    return xlsx_path

# ============================================================
# PDF (one page per product)
# ============================================================

def _product_table(product, pdata, months):
    nm = len(months)
    hdr0 = ["Customer Type", "Product Segment", "Decision"] + sum(
        [[month_title(m), "", ""] for m in months], [])
    hdr1 = ["", "", ""] + ["Rental shortfall", "", "Collection\nRatio"] * nm
    hdr2 = ["", "", ""] + ["No.", "Value (LKR)", ""] * nm
    data = [hdr0, hdr1, hdr2]
    style = list(BASE_TABLE_STYLE) + [
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BLUE),
        ('BACKGROUND', (0, 1), (2, 2), HEADER_BLUE),
        ('BACKGROUND', (3, 1), (-1, 2), SUBHDR_BLUE),
        ('TEXTCOLOR', (0, 0), (-1, 2), colors.white),
        ('FONTNAME', (0, 0), (-1, 2), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, 2), 'CENTER'),
        ('FONTSIZE', (0, 1), (-1, 2), 7.5),
        ('ALIGN', (3, 3), (-1, -1), 'RIGHT'),
        ('ALIGN', (2, 3), (2, -1), 'CENTER'),
    ]
    for col in range(3):
        style.append(('SPAN', (col, 0), (col, 2)))
    for i in range(nm):
        c0 = 3 + 3 * i
        style += [('SPAN', (c0, 0), (c0 + 2, 0)), ('SPAN', (c0, 1), (c0 + 1, 1)),
                  ('SPAN', (c0 + 2, 1), (c0 + 2, 2)),
                  ('LINEAFTER', (c0 + 2, 0), (c0 + 2, -1), 1.0, GRID_COLOR)]

    seg_start = None
    for kind, label, dec, vals in rental_rows(product, pdata, months):
        r = len(data)
        values = [""] * (3 * nm) if vals is None else sum(
            [[format_count(no), format_amount(value), format_ratio(ratio)]
             for no, value, ratio in vals], [])
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

    label_w = [1.3 * inch, 1.6 * inch, 0.7 * inch]
    unit = (CONTENT_WIDTH - sum(label_w)) / (nm * 2.4)
    widths = label_w + [w * unit for w in (0.6, 1.1, 0.7)] * nm
    tbl = Table(data, colWidths=widths, repeatRows=3, hAlign='LEFT')
    tbl.setStyle(TableStyle(style + [('TOPPADDING', (0, 3), (-1, -1), 1.5),
                                     ('BOTTOMPADDING', (0, 3), (-1, -1), 1.5)]))
    return tbl


def create_pdf(target, pdf_path, months, as_at, source_files):
    doc = SimpleDocTemplate(str(pdf_path), pagesize=PAGE_SIZE,
                            leftMargin=PAGE_MARGIN, rightMargin=PAGE_MARGIN,
                            topMargin=PAGE_MARGIN, bottomMargin=PAGE_MARGIN,
                            title=RENTAL_REPORT_TITLE)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('T', parent=styles['Title'], fontSize=16, spaceAfter=2,
                                 alignment=TA_CENTER, fontName='Helvetica-Bold',
                                 textColor=TITLE_NAVY)
    head_style = ParagraphStyle('H', parent=styles['Heading2'], fontSize=12, spaceAfter=2,
                                spaceBefore=2, alignment=TA_LEFT, fontName='Helvetica-Bold',
                                textColor=TABLE_TITLE_BLUE)
    note_style = ParagraphStyle('N', parent=styles['Normal'], fontSize=7, leading=8.5,
                                textColor=colors.grey, spaceAfter=3)
    notes = (f"* No. = number of facilities activated (disbursed) in the month; Value = sum of "
             f"LEASE AMOUNT; Collection Ratio = sum {COLLECTION_COL} / sum {DUE_COL} from that "
             f"month's portfolio. Decision: {CS_DECISION} (Credit Score), else {RBP_DECISION} "
             f"(RBP), else Manual.<br/>* " + customer_rule_text(as_at) + " Source: "
             f"{', '.join(Path(f).name for f in source_files)}.")
    elements = []
    for i, product in enumerate(PRODUCTS):
        if i:
            elements.append(PageBreak())
        pdata = target[target["PRODUCT_CATEGORY"] == product]
        elements += [Paragraph(RENTAL_REPORT_TITLE, title_style),
                     Paragraph(f"<u>{product} (as at {ordinal_date(as_at)})</u> &nbsp; "
                               f"{subtitle(months)}", head_style),
                     Paragraph(notes, note_style),
                     _product_table(product, pdata, months)]
    doc.build(elements)
    log.info(f"PDF saved: {pdf_path}")
    return pdf_path

# ============================================================
# MAIN
# ============================================================

def parse_args():
    ap = argparse.ArgumentParser(description=RENTAL_REPORT_TITLE)
    ap.add_argument("--lending-dir", default=str(LENDING_DIR),
                    help="folder with the monthly LENDING portfolio csvs")
    ap.add_argument("--cs-file", help="Credit Score Decision Report (default: latest)")
    ap.add_argument("--rbp-file", help="RBP Security Approved (default: latest)")
    ap.add_argument("--history-file", help=f"{BDO_PREFIX} history (default: latest multi-month)")
    ap.add_argument("--no-move", action="store_true", help="leave downloads in data/ (testing)")
    return ap.parse_args()


def main():
    args = parse_args()
    setup_logging("rental_shortfall_summary")
    t0 = time.time()
    try:
        # 1. LENDING: the latest 3 months set the report months and the as-at date
        lending, months, _ = load_lending(Path(args.lending_dir))
        as_at = months[-1].end_time.normalize()               # last day of the latest month
        period = f"{months[0].strftime('%B-%Y')} to {months[-1].strftime('%B-%Y')}"
        log.info(f"{RENTAL_REPORT_TITLE} | months {[str(m) for m in months]} | as at {as_at:%d-%b-%Y}")

        # 2. Input files (newest of each, unless given on the command line)
        files = {"CREDIT_SCORE": args.cs_file or find_latest("CREDIT_SCORE"),
                 "RBP": args.rbp_file or find_latest("RBP")}
        try:
            files["HISTORY"] = args.history_file or find_latest("HISTORY")
        except FileNotFoundError:
            log.warning(f"No {BDO_PREFIX} history file found")
        cohort_files = find_cohort_files(months, files.get("HISTORY"))
        for k, f in list(files.items()) + [(f"COHORT {m}", p) for m, p in cohort_files.items()]:
            log.info(f"Input {k:<14}: {Path(f).name}")

        # 3-6. Cohorts, decisions, categories, LENDING merge
        df, target = prepare(lending, months, as_at, cohort_files, files)

        # 7. Metrics (logged per product; the tables are built in step 8)
        for product in PRODUCTS + ["Total"]:
            d = target if product == "Total" else target[target["PRODUCT_CATEGORY"] == product]
            parts = " | ".join(f"{m}: {no:,} / {value:,.0f} / {format_ratio(ratio)}"
                               for m, (no, value, ratio) in zip(months, month_metrics(d, months)))
            log.info(f"Summary {product:<12} (No. / Value / Collection Ratio) {parts}")

        # 8. Outputs: preprocessed data, PDF, Excel; then move the downloads used
        out_dir = REPORTS_DIR / datetime.now().strftime("%Y-%m-%d")
        out_data = out_dir / "data"
        out_data.mkdir(parents=True, exist_ok=True)
        tag = period.replace(" ", "_")
        drop = ["_KEY", "_MONTH"]
        df.drop(columns=drop).to_excel(out_data / f"Rental_Preprocessed_All_{tag}.xlsx", index=False)
        target.drop(columns=drop).to_excel(out_data / f"Rental_Preprocessed_AF_3W_2W_{tag}.xlsx",
                                           index=False)

        pdf_path = out_dir / f"{RENTAL_FILE_NAME}_{tag}.pdf"
        xlsx_path = out_dir / f"{RENTAL_FILE_NAME}_{tag}.xlsx"
        create_pdf(target, pdf_path, months, as_at,
                   list(files.values()) + list(cohort_files.values()))
        create_excel(target, xlsx_path, months, as_at)

        if not args.no_move:
            cohorts_only = {str(m): p for m, p in cohort_files.items()
                            if Path(p) != Path(files.get("HISTORY", ""))}
            move_inputs({**files, **cohorts_only}, out_data)
        log.info(f"===== {RENTAL_REPORT_TITLE} completed in {time.time() - t0:.1f}s =====")

        # Lines read by main_orchestrator.py
        print(f"RECORDS : {len(target)}")
        print(f"AMOUNT  : {target[BDO_AMOUNT_COL].sum():.2f}")
        print(f"PERIOD  : {period}")
        print(f"PDF     : {pdf_path}")
        print(f"EXCEL   : {xlsx_path}")
    except Exception as e:
        log.error(f"Failed to generate {RENTAL_REPORT_TITLE}: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
