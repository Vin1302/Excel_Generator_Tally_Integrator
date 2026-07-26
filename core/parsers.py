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
# -----------------------------------------------------------------------------
# UNIVERSAL HEADER KEYWORDS
# -----------------------------------------------------------------------------

_KW_DATE = (
    "date",
    "txn date",
    "transaction date",
    "value date",
    "post date",
    "posting date",
)

_KW_DESC = (
    "remarks",
    "remark",
    "narration",
    "particular",
    "particulars",
    "description",
    "details",
    "detail",
    "transaction details",
)

_KW_REFERENCE = (
    "ref",
    "reference",
    "ref no",
    "cheque",
    "cheque no",
    "instrument",
    "instrument no",
)

_KW_DEBIT = (
    "withdrawal",
    "withdraw",
    "withdrawal amt",
    "debit",
    "dr",
    "paid out",
)

_KW_CREDIT = (
    "deposit",
    "credit",
    "cr",
    "paid in",
)

_KW_BALANCE = (
    "balance",
    "closing balance",
    "running balance",
    "available balance",
)

_KW_AMOUNT = (
    "amount",
    "txn amount",
    "transaction amount",
    "amt",
)

# Header abbreviations seen in bank-statement PDFs.  These are deliberately
# kept separate from the general keywords above: a lone "dr" or "cr" is only
# meaningful when locating column headings, not when interpreting narration.
_HEADER_DEBIT_ABBRS = {"dr", "drs"}
_HEADER_CREDIT_ABBRS = {"cr", "crs"}

# Keywords used only to locate the header ROW inside messy Excel/PDF sheets.
_HEADER_ROW_KEYWORDS = (
    "date", "remark", "narration", "particular", "description", "detail",
    "withdrawal", "debit", "deposit", "credit", "amount", "balance",
)


def _norm_header(h) -> str:
    return re.sub(r"\s+", "", str(h).strip().lower())


def _detect_roles(df: pd.DataFrame):

    roles = {}

    for col in df.columns:

        key = _norm_header(col)

        key = (
            key.replace("₹", "")
               .replace("rs.", "")
               .replace("rs", "")
               .replace("_", "")
               .strip()
        )

        # ---------------- DATE ----------------

        if "valuedate" in key:
            roles["date"] = col
            continue

        if "transactiondate" in key:
            roles.setdefault("date", col)
            continue

        if "postdate" in key:
            roles.setdefault("post_date", col)
            continue

        if any(x.replace(" ","") in key for x in _KW_DATE):
            roles.setdefault("date", col)
            continue

        # ---------------- DESCRIPTION ----------------

        if any(x.replace(" ","") in key for x in _KW_DESC):
            roles.setdefault("description", col)
            continue

        # ---------------- REFERENCE ----------------

        if any(x.replace(" ","") in key for x in _KW_REFERENCE):
            roles.setdefault("reference", col)
            continue

        # ---------------- DEBIT ----------------

        if any(x.replace(" ","") in key for x in _KW_DEBIT):
            roles.setdefault("debit", col)
            continue

        # ---------------- CREDIT ----------------

        if any(x.replace(" ","") in key for x in _KW_CREDIT):
            roles.setdefault("credit", col)
            continue

        # ---------------- BALANCE ----------------

        if any(x.replace(" ","") in key for x in _KW_BALANCE):
            roles.setdefault("balance", col)
            continue

        # ---------------- AMOUNT ----------------

        if any(x.replace(" ","") in key for x in _KW_AMOUNT):
            roles.setdefault("amount", col)

    return roles

def _to_number(value):

    if value is None:
        return 0.0

    if pd.isna(value):
        return 0.0

    s = str(value).strip()

    if s == "":
        return 0.0

    if s.lower() in {"nan", "none", "-", "--"}:
        return 0.0

    # Remove common currency markers
    s = (
        s.replace("₹", "")
         .replace("Rs.", "")
         .replace("Rs", "")
         .replace("INR", "")
         .replace(",", "")
         .strip()
    )

    negative = False

    # Parentheses indicate a negative amount
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()

    lower = s.lower()

    if lower.endswith("dr"):
        negative = True
        s = s[:-2].strip()

    elif lower.endswith("cr"):
        s = s[:-2].strip()

    # Remove anything that's not part of a number
    s = re.sub(r"[^0-9.\-]", "", s)

    if s in {"", ".", "-"}:
        return 0.0

    try:
        value = float(s)
    except ValueError:
        return 0.0

    return -abs(value) if negative else value

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
    desc = str(row.get(roles["description"], "")).strip()

if roles.get("reference") is not None:
    ref = str(row.get(roles["reference"], "")).strip()

    if ref and ref.lower() != "nan":
        desc += " " + ref

desc = re.sub(r"\s+", " ", desc).strip()

payee = desc

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

# Date tokens: dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy, dd-MON-yy, yyyy-mm-dd,
# and the ``dd Mon yyyy`` form used by SBI account-statement PDFs.
_DATE_RE = re.compile(
    r"\b(\d{1,2}[./-][A-Za-z0-9]{2,4}[./-]\d{2,4}|"
    r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}|\d{4}-\d{2}-\d{2})\b",
    re.IGNORECASE,
)
_AMOUNT_RE = re.compile(r"-?\(?\d[\d,]*\.\d{2}\)?")


def _pdf_page_via_ocr(page, source):
    """Last resort: use EasyOCR to extract text from image-only PDF pages.

    Pure Python OCR - no system dependencies required!
    Works offline after first download of language model.
    """
    try:
        import easyocr
    except ImportError:
        log.debug("easyocr not installed; skipping OCR")
        return _empty()

    try:
        # Import fitz for page-to-image conversion
        import fitz
        import numpy as np
    except ImportError:
        log.debug("fitz/numpy not available for OCR")
        return _empty()

    try:
        log.info("Converting PDF page to image for OCR...")
        # Convert PDF page to image with good resolution
        mat = fitz.Matrix(2, 2)  # 2x zoom for better OCR
        pix = page.get_pixmap(matrix=mat, alpha=False)

        # Convert pixmap to numpy array (BGR format for OpenCV)
        img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

        # Initialize OCR reader (English - downloads ~200MB model on first run)
        log.info("Initializing OCR reader (first run may take a moment)...")
        reader = easyocr.Reader(['en'], gpu=False, verbose=False)

        # Run OCR
        log.info("Running OCR extraction...")
        results = reader.readtext(img_array)

        if not results:
            log.debug("OCR produced no text")
            return _empty()

        # Extract text from results
        ocr_text = "\n".join([text[1] for text in results])
        log.info("OCR extracted %d chars from page", len(ocr_text))

        # Parse the OCR'd text using an enhanced method that handles
        # dates and amounts on separate lines
        rows = []
        lines = ocr_text.splitlines()

        i = 0
        while i < len(lines):
            line = lines[i]
            dm = _DATE_RE.search(line)

            # If we found a date, look for amounts in the next few lines
            if dm:
                date = pd.to_datetime(dm.group(1), errors="coerce", dayfirst=True)
                if not pd.isna(date):
                    # Collect amounts from this line and next 3 lines
                    desc_combined = line
                    amounts_found = []

                    # Look in current line
                    amounts_in_line = _AMOUNT_RE.findall(line)
                    amounts_found.extend(amounts_in_line)

                    # Look in next 3 lines for amounts
                    for j in range(i + 1, min(i + 4, len(lines))):
                        next_line = lines[j]
                        amounts_in_next = _AMOUNT_RE.findall(next_line)
                        if amounts_in_next:
                            amounts_found.extend(amounts_in_next)
                            desc_combined += " " + next_line
                        # Stop if we hit another date
                        elif _DATE_RE.search(next_line):
                            break

                    # If we found amount(s), use the last one (usually the transaction amount)
                    if amounts_found:
                        amt = _to_number(amounts_found[-1])
                        if amt != 0:
                            # Clean up description
                            desc = re.sub(r"\s{2,}", " ", desc_combined).strip(" .-")
                            # Remove dates and amounts from description
                            desc = _DATE_RE.sub("", desc)
                            for a in amounts_found:
                                desc = desc.replace(a, "")
                            desc = re.sub(r"\s{2,}", " ", desc).strip(" .-")

                            # Determine direction
                            if "WDL" in line or "withdrawal" in line.lower() or "debit" in line.lower():
                                direction = "paid"
                            elif "DEP" in line or "deposit" in line.lower() or "credit" in line.lower():
                                direction = "received"
                            else:
                                # Guess from sign
                                direction = "paid" if amt < 0 else "received"

                            if desc.strip():  # Only add if description exists
                                rows.append({
                                    "date": date,
                                    "description": desc,
                                    "payee": desc,
                                    "amount": abs(amt),
                                    "direction": direction,
                                    "balance": None,
                                    "source": source
                                })

            i += 1

        return pd.DataFrame(rows, columns=_COLUMNS) if rows else _empty()

    except Exception as e:
        log.debug("OCR extraction failed: %s", e)
        return _empty()


def parse_pdf(path: str) -> pd.DataFrame:
    source = os.path.basename(path)
    try:
        import pdfplumber
    except ImportError:
        log.error("pdfplumber is not installed; cannot read PDFs.")
        return _empty()

    all_norm: List[pd.DataFrame] = []
    anchors_holder = {"anchors": None}   # header positions, remembered across pages
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            # 1) Position-aware reader (best for Withdrawal/Deposit/Balance
            #    layouts like ICICI, HDFC, SBI, Axis savings statements).
            df = _pdf_page_via_positions(page, source, anchors_holder)
            # 2) Fall back to generic table extraction.
            if df is None or df.empty:
                df = _pdf_page_via_tables(page, source)
            # 3) Last resort: line-by-line text.
            if df is None or df.empty:
                df = _pdf_page_via_text(page, source)
            # 4) OCR fallback: use Tesseract for image-only pages.
            if df is None or df.empty:
                df = _pdf_page_via_ocr(page, source)
            if df is not None and not df.empty:
                all_norm.append(df)

    if not all_norm:
        log.warning("%s: no transactions extracted from PDF", source)
        return _empty()
    return pd.concat(all_norm, ignore_index=True)


# --------------------------------------------------------------------------- #
# Position-aware extractor
# --------------------------------------------------------------------------- #
_MONEY_TOKEN_RE = re.compile(r"^-?\(?\d[\d,]*\.\d{2}\)?$")
_DATE_TOKEN_RE = re.compile(
    r"^(?:\d{1,2}[./-][A-Za-z0-9]{2,4}[./-]\d{2,4}|"
    r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}|\d{4}-\d{2}-\d{2})$",
    re.IGNORECASE,
)


def _is_money_token(tok: str) -> bool:
    return bool(_MONEY_TOKEN_RE.match(str(tok).strip()))


def _center(w) -> float:
    return (w["x0"] + w["x1"]) / 2.0


def _group_words_into_lines(words, tol: float = 3.0):
    """Group extracted words into visual lines by their vertical position."""
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, cur, cur_top = [], [], None
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= tol:
            cur.append(w)
            cur_top = w["top"] if cur_top is None else cur_top
        else:
            lines.append(sorted(cur, key=lambda x: x["x0"]))
            cur, cur_top = [w], w["top"]
    if cur:
        lines.append(sorted(cur, key=lambda x: x["x0"]))
    return lines


def _amount_header_role(text: str):
    """Return the amount-column role named by one PDF header word, if any."""
    key = re.sub(r"[^a-z]", "", str(text).lower())
    if not key:
        return None
    if key in _HEADER_DEBIT_ABBRS or any(kw in key for kw in _KW_DEBIT):
        return "withdrawal"
    if key in _HEADER_CREDIT_ABBRS or any(kw in key for kw in _KW_CREDIT):
        return "deposit"
    if "balance" in key or key in {"bal", "balrs"}:
        return "balance"
    return None

def _detect_pdf_columns(lines):
    """
    Detect every visible PDF column by reading the header text.
    Returns

    {
        "date":120,
        "post_date":180,
        "description":250,
        "reference":420,
        "debit":620,
        "credit":760,
        "balance":900
    }
    """

    for start in range(min(len(lines), 12)):

        words = []

        # SBI headers often wrap over several lines
        for ln in lines[start:start+6]:
            words.extend(ln)

        cols = {}

        for w in words:

            txt = (
                w["text"]
                .lower()
                .replace("₹","")
                .replace("rs.","")
                .replace("rs","")
                .strip()
            )

            x = w["x0"]

            if "value" in txt and "date" in txt:
                cols["date"] = x

            elif "transaction" in txt and "date" in txt:
                cols.setdefault("date",x)

            elif "post" in txt:
                cols["post_date"] = x

            elif txt in {
                "details",
                "detail",
                "remarks",
                "description",
                "narration",
                "particulars"
            }:
                cols["description"] = x

            elif "ref" in txt or "cheque" in txt:
                cols["reference"] = x

            elif txt in {
                "withdrawal",
                "debit",
                "dr"
            }:
                cols["debit"] = x

            elif txt in {
                "deposit",
                "credit",
                "cr"
            }:
                cols["credit"] = x

            elif "balance" in txt:
                cols["balance"] = x

        if "date" in cols and "balance" in cols:
            return cols

    return None

def _find_amount_anchors(lines):
    """Locate Withdrawal/Debit, Deposit/Credit and Balance column headings.

    PDF producers often wrap SBI's header over several visual lines (for
    example, ``Debit Credit`` on one line and ``Balance`` below it).  Search a
    small adjacent-line window rather than requiring all labels on one line.
    """
    for start in range(len(lines)):
        words = [w for ln in lines[start:start + 3] for w in ln]
        roles = {}
        for w in words:
            role = _amount_header_role(w["text"])
            if role:
                # A heading can be split into multiple words; its left-most
                # word is the correct column boundary.
                roles[role] = min(roles.get(role, w["x0"]), w["x0"])

        if {"withdrawal", "deposit", "balance"} <= set(roles):
            # The layouts handled by this coordinate reader place debit,
            # credit and balance from left to right.  Reject narration that
            # happens to mention the same words in another order.
            if roles["withdrawal"] < roles["deposit"] < roles["balance"]:
                return roles
    return None


def _find_column(x, anchors):
    """
    Returns which detected column an x-coordinate belongs to.
    """

    ordered = sorted(
        anchors.items(),
        key=lambda x: x[1]
    )

    previous = ordered[0][0]

    for name,pos in ordered:

        if x < pos:
            return previous

        previous = name

    return ordered[-1][0]


def _pdf_page_via_positions(page, source, holder):
    """Read a statement page using word coordinates. Requires a header row with
    Withdrawal/Deposit/Balance columns (found on this or an earlier page)."""
    try:
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    except Exception:
        return _empty()
    if not words:
        return _empty()

    lines = _group_words_into_lines(words)

   anchors = _detect_pdf_columns(lines)
    if anchors:
        holder["anchors"] = anchors
    anchors = holder.get("anchors")
    if not anchors:
        return _empty()   # no Withdrawal/Deposit/Balance layout -> let fallbacks try

    records = []
    current = None
    for ln in lines:
        joined = " ".join(w["text"] for w in ln)
        lower = joined.lower()

if (
    "value date" in lower
    or "post date" in lower
    or "details" in lower
    or "description" in lower
    or "withdrawal" in lower
    or "debit" in lower
    or "credit" in lower
    or "balance" in lower
):
    continue
        dm = _DATE_RE.search(joined)
        moneys = [(w, _center(w)) for w in ln if _is_money_token(w["text"])]

        is_start = bool(dm) and bool(moneys)
        if is_start:
            if current:
                records.append(current)
            wd = dep = bal = None
            for w, xc in moneys:
                role = _find_column(xc, anchors)
                val = _to_number(w["text"])
              if role in ("withdrawal", "debit"):
    wd = val

elif role in ("deposit", "credit"):
    dep = val

elif role == "balance":
    bal = val
            # First-line remarks: words left of the Withdrawal column, minus the
            # leading serial number and the date token.
          desc_words = []
ref_words = []

for w in ln:
    x = _center(w)

    # Description column
    if anchors.get("description") and anchors.get("debit"):
        if anchors["description"] <= x < anchors["debit"]:
            desc_words.append(w["text"])

    # Reference column (if present)
    if anchors.get("reference") and anchors.get("debit"):
        if anchors["reference"] <= x < anchors["debit"]:
            ref_words.append(w["text"])

remarks = " ".join(desc_words)

if ref_words:
    remarks += " " + " ".join(ref_words)

remarks = _DATE_RE.sub("", remarks)
remarks = re.sub(r"\s+", " ", remarks).strip()
            date = pd.to_datetime(dm.group(1), errors="coerce", dayfirst=True)
            current = {"date": date, "payee": remarks, "remarks": remarks,
                       "wd": wd, "dep": dep, "bal": bal}
       elif current is not None:
            # Continuation line (wrapped UPI/reference text) -> append to remarks.
            cont = []

for w in ln:
    x = _center(w)

    # Only append words from the narration area.
    if (
        anchors.get("description") is not None
        and anchors.get("debit") is not None
        and anchors["description"] <= x < anchors["debit"]
    ):
        cont.append(w["text"])

if cont:
    current["remarks"] += " " + " ".join(cont)
    if current:
        records.append(current)

    out = []
    for r in records:
        if pd.isna(r["date"]):
            continue
        if r["wd"] and abs(r["wd"]) > 0:
            amount, direction = abs(r["wd"]), "paid"
        elif r["dep"] and abs(r["dep"]) > 0:
            amount, direction = abs(r["dep"]), "received"
        else:
            continue
        desc = re.sub(r"\s{2,}", " ", r["remarks"]).strip()
        out.append({
            "date": r["date"], "description": desc,
            "payee": r["payee"] or desc, "amount": amount,
            "direction": direction, "balance": r["bal"], "source": source,
        })
    return pd.DataFrame(out, columns=_COLUMNS) if out else _empty()


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
    """Last-resort fallback: scan text lines. Grab the first date and the LAST
    money-looking number on each line; text in between becomes the description."""
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
