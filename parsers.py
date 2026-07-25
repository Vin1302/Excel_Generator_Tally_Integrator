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
# Header recognition. We match by KEYWORDS CONTAINED IN the header (after
# lower-casing and stripping spaces), so real-world names like
# "Withdrawal Amount (INR)" or "Transaction Remarks" are recognised without
# needing an exact match. Order of checks matters (see _detect_roles).
# --------------------------------------------------------------------------- #
_KW_DESC = ("remark", "narration", "particular", "description", "detail")
_KW_DEBIT = ("withdrawal", "debit", "paidout")
_KW_CREDIT = ("deposit", "credit", "paidin")
_KW_TYPE = ("drcr", "crdr", "indicator")

# Keywords used only to locate the header ROW inside messy Excel/PDF sheets.
_HEADER_ROW_KEYWORDS = (
    "date", "remark", "narration", "particular", "description", "detail",
    "withdrawal", "debit", "deposit", "credit", "amount", "balance",
)


def _norm_header(h) -> str:
    return re.sub(r"\s+", "", str(h).strip().lower())


def _detect_roles(df: pd.DataFrame) -> dict:
    """Assign each column a role (date / description / debit / credit / type /
    amount) by looking for keywords inside the header text. A 'balance' column
    is explicitly ignored so it's never mistaken for a transaction amount."""
    roles = {}
    for col in df.columns:
        key = _norm_header(col)
        if not key:
            continue
        if "balance" in key:
            roles.setdefault("balance", col)   # kept for the as-is sheet, never
            continue                            # treated as a transaction amount
        if "date" in key:
            roles.setdefault("date", col)
        elif any(w in key for w in _KW_DESC):
            roles.setdefault("description", col)
        elif any(w in key for w in _KW_DEBIT):        # before generic "amount"
            roles.setdefault("debit", col)
        elif any(w in key for w in _KW_CREDIT):       # before generic "amount"
            roles.setdefault("credit", col)
        elif "type" in key or any(w in key for w in _KW_TYPE):
            roles.setdefault("type", col)
        elif "amount" in key or key in {"amt", "value"}:
            roles.setdefault("amount", col)
    return roles


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
def _dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Make column names unique. Real bank files often have blank or repeated
    headers (e.g. two 'Amount' columns, or several empty ones); without this,
    df[col] can return a whole DataFrame and pandas raises errors like
    'cannot assemble with duplicate keys'."""
    seen = {}
    new_cols = []
    for i, col in enumerate(df.columns):
        name = "" if col is None else str(col).strip()
        if name == "":
            name = f"col_{i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        new_cols.append(name)
    df = df.copy()
    df.columns = new_cols
    return df


def _dataframe_to_normalized(df: pd.DataFrame, source: str) -> pd.DataFrame:
    """Map an arbitrary bank table (already read into a DataFrame with a
    header row) onto the normalized schema."""
    if df is None or df.empty:
        return _empty()

    df = _dedupe_columns(df)   # <-- guarantees unique column labels

    roles = _detect_roles(df)

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
        payee = ""
        if roles.get("description") is not None:
            raw_desc = str(row.get(roles["description"], "") or "").strip()
            lines = [ln.strip() for ln in raw_desc.splitlines() if ln.strip()]
            # payee   = first line (the counterparty) -> used for grouping
            # desc    = full remark, newlines flattened -> shown as-is
            payee = lines[0] if lines else raw_desc
            desc = " ".join(lines) if lines else raw_desc

        amount, direction = _resolve_amount(row, roles)
        if amount == 0:
            continue

        balance = None
        if roles.get("balance") is not None:
            b = _to_number(row.get(roles["balance"]))
            balance = b if b != 0 else None

        out_rows.append(
            {
                "date": date,
                "description": desc,
                "payee": payee or desc,
                "amount": abs(amount),
                "direction": direction,
                "balance": balance,
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
    for i in range(min(len(raw), 20)):
        row_text = " ".join(_norm_header(c) for c in raw.iloc[i].tolist())
        hits = sum(1 for kw in _HEADER_ROW_KEYWORDS if kw in row_text)
        if hits >= 2:   # a real header row mentions at least a couple of these
            return i
    return 0


# --------------------------------------------------------------------------- #
# PDF
# --------------------------------------------------------------------------- #
# We try each of these table-extraction strategies and keep whichever pulls the
# most transactions. "lines" suits statements with ruled borders; "text" suits
# borderless ones that align columns by whitespace. Trying all makes the reader
# work across many bank layouts without manual tuning.
_PDF_TABLE_STRATEGIES = [
    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
    {"vertical_strategy": "text", "horizontal_strategy": "text"},
    {"vertical_strategy": "lines", "horizontal_strategy": "text"},
]

# Date tokens: dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy, dd-MON-yy, yyyy-mm-dd, etc.
_DATE_RE = re.compile(
    r"\b(\d{1,2}[./-][A-Za-z0-9]{2,4}[./-]\d{2,4}|\d{4}-\d{2}-\d{2})\b"
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
    best = _empty()
    for settings in _PDF_TABLE_STRATEGIES:
        try:
            tables = page.extract_tables(settings)
        except Exception:
            continue
        for table in tables or []:
            if not table or len(table) < 2:
                continue
            header = [str(c or "") for c in table[0]]
            body = table[1:]
            try:
                df = pd.DataFrame(body, columns=header)
            except Exception:
                continue
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
            {"date": date, "description": desc, "payee": desc,
             "amount": abs(amt), "direction": direction,
             "balance": None, "source": source}
        )
    return pd.DataFrame(rows, columns=_COLUMNS) if rows else _empty()


# --------------------------------------------------------------------------- #
# Dispatcher
# --------------------------------------------------------------------------- #
_COLUMNS = ["date", "description", "payee", "amount", "direction", "balance", "source"]


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
