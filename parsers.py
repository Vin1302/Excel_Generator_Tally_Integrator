"""
parsers.py
----------
Read bank-statement files (CSV, Excel, PDF) and turn each into a *normalized*
pandas DataFrame with these exact columns:

    date         : datetime64  (the transaction date)
    description  : str         (raw narration / particulars, uncleaned)
    amount       : float       (always POSITIVE)
    direction    : str         ('paid'  = money out,  'received' = money in)
    source       : str         (the file it came from)

Everything downstream (grouping, Excel report, Tally export) depends only on
these five columns, so if you ever add a new file type you just need to emit
this same shape.

NOTE ON PDFs: every bank lays out its PDF differently. The extractor here is a
reasonable general-purpose attempt (table extraction first, line-by-line text
as a fallback). For a specific bank you may need to tweak `_pdf_table_settings`
or the fallback regex. Those two spots are commented clearly.
"""

from __future__ import annotations

import os
import re
import logging
import warnings
from typing import List

import pandas as pd

log = logging.getLogger(__name__)

# Bank dates come in many formats; we intentionally let pandas fall back to the
# flexible parser, so silence its "could not infer format" chatter.
warnings.filterwarnings("ignore", message="Could not infer format")

# --------------------------------------------------------------------------- #
# Header aliases: how we recognise columns regardless of what the bank calls
# them. All comparisons are done in lower-case with spaces stripped.
# --------------------------------------------------------------------------- #
DATE_HEADERS = {
    "date", "txndate", "transactiondate", "valuedate", "postingdate",
    "bookingdate", "date(valuedate)", "trandate",
}
DESC_HEADERS = {
    "description", "narration", "particulars", "details", "remarks",
    "transactiondetails", "transactionremarks", "chequeref", "reference",
    "transactiondescription",
}
DEBIT_HEADERS = {
    "debit", "withdrawal", "withdrawalamt", "withdrawalamt.", "dr",
    "paidout", "debitamount", "withdrawals", "amountdebit", "debit(dr)",
}
CREDIT_HEADERS = {
    "credit", "deposit", "depositamt", "depositamt.", "cr",
    "paidin", "creditamount", "deposits", "amountcredit", "credit(cr)",
}
AMOUNT_HEADERS = {"amount", "value", "transactionamount", "amt"}
TYPE_HEADERS = {"type", "drcr", "dr/cr", "crdr", "transactiontype", "indicator"}


def _norm_header(h) -> str:
    return re.sub(r"\s+", "", str(h).strip().lower())


def _to_number(x) -> float:
    """Parse a money value that may contain commas, currency symbols,
    parentheses (negative) or trailing Dr/Cr markers. Blank -> 0.0."""
    if x is None:
        return 0.0
    s = str(x).strip()
    if s == "" or s.lower() in {"nan", "none", "-"}:
        return 0.0
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    if s.lower().endswith("dr"):
        negative = True
        s = s[:-2]
    elif s.lower().endswith("cr"):
        s = s[:-2]
    s = re.sub(r"[^0-9.\-]", "", s)   # drop currency symbols, commas, spaces
    if s in {"", ".", "-"}:
        return 0.0
    try:
        val = float(s)
    except ValueError:
        return 0.0
    return -abs(val) if negative else val


# --------------------------------------------------------------------------- #
# Shared: a table-like DataFrame -> normalized rows
# --------------------------------------------------------------------------- #
def _dataframe_to_normalized(df: pd.DataFrame, source: str) -> pd.DataFrame:
    """Map an arbitrary bank table (already read into a DataFrame with a
    header row) onto the normalized schema."""
    if df is None or df.empty:
        return _empty()

    # Map each real column to a role using the alias sets.
    roles = {}
    for col in df.columns:
        key = _norm_header(col)
        if key in DATE_HEADERS:
            roles.setdefault("date", col)
        elif key in DESC_HEADERS:
            roles.setdefault("description", col)
        elif key in DEBIT_HEADERS:
            roles.setdefault("debit", col)
        elif key in CREDIT_HEADERS:
            roles.setdefault("credit", col)
        elif key in AMOUNT_HEADERS:
            roles.setdefault("amount", col)
        elif key in TYPE_HEADERS:
            roles.setdefault("type", col)

    # If we couldn't find a date column by name, guess the most date-like one.
    if "date" not in roles:
        roles["date"] = _guess_date_column(df)
    # If no description column by name, guess the widest text column.
    if "description" not in roles:
        roles["description"] = _guess_text_column(df, exclude=set(roles.values()))

    if roles.get("date") is None:
        log.debug("%s: no date column in this table candidate; skipping it", source)
        return _empty()

    out_rows = []
    for _, row in df.iterrows():
        raw_date = row.get(roles["date"])
        date = pd.to_datetime(raw_date, errors="coerce", dayfirst=True)
        if pd.isna(date):
            continue  # header leftovers / blank / total lines

        desc = ""
        if roles.get("description") is not None:
            desc = str(row.get(roles["description"], "") or "").strip()

        amount, direction = _resolve_amount(row, roles)
        if amount == 0:
            continue

        out_rows.append(
            {
                "date": date,
                "description": desc,
                "amount": abs(amount),
                "direction": direction,
                "source": source,
            }
        )

    return pd.DataFrame(out_rows, columns=_COLUMNS) if out_rows else _empty()


def _resolve_amount(row, roles):
    """Return (amount, direction) for a single row given the detected roles."""
    # Case 1: explicit Debit / Credit columns.
    if "debit" in roles or "credit" in roles:
        debit = _to_number(row.get(roles["debit"])) if "debit" in roles else 0.0
        credit = _to_number(row.get(roles["credit"])) if "credit" in roles else 0.0
        if abs(debit) > 0:
            return abs(debit), "paid"
        if abs(credit) > 0:
            return abs(credit), "received"
        return 0.0, "paid"

    # Case 2: single amount column, maybe with a Dr/Cr type column.
    if "amount" in roles:
        amt = _to_number(row.get(roles["amount"]))
        if "type" in roles:
            t = str(row.get(roles["type"], "")).strip().lower()
            if t.startswith("d") or "debit" in t or t == "dr":
                return abs(amt), "paid"
            if t.startswith("c") or "credit" in t or t == "cr":
                return abs(amt), "received"
        # No type column: rely on sign (negative = money out).
        if amt < 0:
            return abs(amt), "paid"
        return abs(amt), "received"

    return 0.0, "paid"


def _guess_date_column(df: pd.DataFrame):
    best, best_score = None, 0
    for col in df.columns:
        parsed = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        score = parsed.notna().sum()
        if score > best_score:
            best, best_score = col, score
    # Require at least a couple of parseable dates to accept the guess.
    return best if best_score >= 2 else None


def _guess_text_column(df: pd.DataFrame, exclude=frozenset()):
    best, best_len = None, -1
    for col in df.columns:
        if col in exclude:
            continue
        try:
            avg_len = df[col].astype(str).str.len().mean()
        except Exception:
            continue
        if avg_len is not None and avg_len > best_len:
            best, best_len = col, avg_len
    return best


# --------------------------------------------------------------------------- #
# CSV / Excel
# --------------------------------------------------------------------------- #
def parse_csv(path: str) -> pd.DataFrame:
    source = os.path.basename(path)
    # Some bank CSVs have junk (logo/address) lines above the real header.
    # Try reading a few candidate header offsets and keep the best one.
    best = _empty()
    for skip in range(0, 15):
        try:
            df = pd.read_csv(path, skiprows=skip, dtype=str,
                             keep_default_na=False, engine="python")
        except Exception:
            continue
        if df.shape[1] < 2:
            continue
        norm = _dataframe_to_normalized(df, source)
        if len(norm) > len(best):
            best = norm
    return best


def parse_excel(path: str) -> pd.DataFrame:
    source = os.path.basename(path)
    best = _empty()
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
    except Exception as exc:
        log.warning("%s: could not open Excel file (%s)", source, exc)
        return best

    for _, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        # Find the row that looks most like a header (contains date/desc words).
        header_idx = _find_header_row(raw)
        df = raw.iloc[header_idx + 1:].copy()
        df.columns = [str(c) for c in raw.iloc[header_idx].tolist()]
        norm = _dataframe_to_normalized(df, source)
        if len(norm) > len(best):
            best = norm
    return best


def _find_header_row(raw: pd.DataFrame) -> int:
    keywords = DATE_HEADERS | DESC_HEADERS | DEBIT_HEADERS | CREDIT_HEADERS | AMOUNT_HEADERS
    for i in range(min(len(raw), 20)):
        cells = {_norm_header(c) for c in raw.iloc[i].tolist()}
        if cells & keywords:
            return i
    return 0


# --------------------------------------------------------------------------- #
# PDF
# --------------------------------------------------------------------------- #
# Tune this if a particular bank's tables don't extract cleanly. "lines" works
# when the statement has ruled table borders; "text" works when it doesn't.
_pdf_table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
}

_DATE_RE = re.compile(
    r"\b(\d{1,2}[/-][A-Za-z0-9]{2,4}[/-]\d{2,4}|\d{4}-\d{2}-\d{2})\b"
)
_AMOUNT_RE = re.compile(r"-?\(?\d[\d,]*\.\d{2}\)?")


def parse_pdf(path: str) -> pd.DataFrame:
    source = os.path.basename(path)
    try:
        import pdfplumber
    except ImportError:
        log.error("pdfplumber is not installed; cannot read PDFs.")
        return _empty()

    all_norm: List[pd.DataFrame] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            df = _pdf_page_via_tables(page, source)
            if df is None or df.empty:
                df = _pdf_page_via_text(page, source)
            if df is not None and not df.empty:
                all_norm.append(df)

    if not all_norm:
        log.warning("%s: no transactions extracted from PDF", source)
        return _empty()
    return pd.concat(all_norm, ignore_index=True)


def _pdf_page_via_tables(page, source):
    try:
        tables = page.extract_tables(_pdf_table_settings)
    except Exception:
        tables = []
    best = _empty()
    for table in tables or []:
        if not table or len(table) < 2:
            continue
        header = [str(c or "") for c in table[0]]
        body = table[1:]
        df = pd.DataFrame(body, columns=header)
        norm = _dataframe_to_normalized(df, source)
        if len(norm) > len(best):
            best = norm
    return best


def _pdf_page_via_text(page, source):
    """Fallback: scan text lines. Grab the first date and the LAST money-looking
    number on each line; text in between becomes the description. Direction is
    guessed from sign only, so verify against a real statement and adjust."""
    try:
        text = page.extract_text() or ""
    except Exception:
        return _empty()

    rows = []
    for line in text.splitlines():
        dm = _DATE_RE.search(line)
        if not dm:
            continue
        amounts = _AMOUNT_RE.findall(line)
        if not amounts:
            continue
        date = pd.to_datetime(dm.group(1), errors="coerce", dayfirst=True)
        if pd.isna(date):
            continue
        amt = _to_number(amounts[-1])
        if amt == 0:
            continue
        desc = line[dm.end():].strip()
        for a in amounts:
            desc = desc.replace(a, "")
        desc = re.sub(r"\s{2,}", " ", desc).strip(" .-")
        direction = "paid" if amt < 0 else "received"
        rows.append(
            {"date": date, "description": desc, "amount": abs(amt),
             "direction": direction, "source": source}
        )
    return pd.DataFrame(rows, columns=_COLUMNS) if rows else _empty()


# --------------------------------------------------------------------------- #
# Dispatcher
# --------------------------------------------------------------------------- #
_COLUMNS = ["date", "description", "amount", "direction", "source"]


def _empty() -> pd.DataFrame:
    return pd.DataFrame(columns=_COLUMNS)


def parse_file(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv" or ext == ".tsv":
        return parse_csv(path)
    if ext in {".xlsx", ".xlsm", ".xls", ".xltx"}:
        return parse_excel(path)
    if ext == ".pdf":
        return parse_pdf(path)
    log.warning("Unsupported file type: %s", path)
    return _empty()


def parse_many(paths: List[str]) -> pd.DataFrame:
    frames = [parse_file(p) for p in paths]
    frames = [f for f in frames if f is not None and not f.empty]
    if not frames:
        return _empty()
    combined = pd.concat(frames, ignore_index=True)
    combined = combined.sort_values("date").reset_index(drop=True)
    return combined
