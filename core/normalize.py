"""
normalize.py
------------
Turn messy transaction descriptions into tidy *groups*, and map each group to a
Tally ledger name.

Two jobs:

1. clean_name()          -> strips card numbers, reference IDs, dates and common
                            noise words so that "AMZN Mktp UK*A1B2" and
                            "AMAZON PRIME" both reduce to something comparable.

2. group_transactions()  -> fuzzy-clusters the cleaned names so that near-identical
                            payees collapse into one group (uses rapidfuzz).

3. load_ledger_map() /   -> map a group's canonical name to the Tally ledger it
   attach_ledgers()         should post to (editable CSV, grows over time).
"""

from __future__ import annotations

import os
import re
import logging

import pandas as pd

log = logging.getLogger(__name__)

# Tokens that carry no identity and just add noise to payee names.
_NOISE_WORDS = {
    "POS", "UPI", "IMPS", "NEFT", "RTGS", "ACH", "ATM", "TFR", "TRANSFER",
    "PAYMENT", "PMT", "REF", "REFNO", "TXN", "TRANSACTION", "PURCHASE",
    "DEBIT", "CREDIT", "CARD", "VISA", "MASTERCARD", "RRN", "BIL", "BILL",
    "ONLINE", "MOBILE", "BANKING", "TO", "FROM", "THE", "LTD", "PVT",
    "INDIA", "IN", "GBP", "USD", "INR", "WWW", "COM",
}


def clean_name(raw: str) -> str:
    """Reduce a raw narration to a comparable payee name."""
    if not raw:
        return "UNKNOWN"
    # If a multi-line remark slipped through, the payee is the first line.
    first_line = str(raw).splitlines()[0] if str(raw).strip() else str(raw)
    s = first_line.upper()

    # Remove dates like 12/03/2024, 2024-03-12, 12-MAR-24.
    s = re.sub(r"\b\d{1,2}[/-][A-Z0-9]{2,4}[/-]\d{2,4}\b", " ", s)
    # Remove long digit runs (card/account/reference numbers), e.g. the 687972
    # in "bajajpay.687972". Short trailing digits in a handle are kept so real
    # payees like "siyabodadkar53" survive.
    s = re.sub(r"\d{4,}", " ", s)
    # Replace any non-letter/number with a space.
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    # Drop noise words.
    tokens = [t for t in s.split() if t not in _NOISE_WORDS and len(t) > 1]
    cleaned = " ".join(tokens).strip()

    if not cleaned:
        return "UNKNOWN"
    # Keep it short-ish: the leading words usually carry the identity.
    return " ".join(cleaned.split()[:4])


def group_transactions(df: pd.DataFrame, threshold: int = 88) -> pd.DataFrame:
    """Add two columns to the transaction frame:

        clean_name : the cleaned payee name (per row)
        group      : the canonical name shared by all fuzzily-matching rows

    Grouping is done separately within 'paid' and 'received' so an incoming and
    outgoing payment to the same party don't merge.
    """
    if df.empty:
        df = df.copy()
        df["clean_name"] = []
        df["group"] = []
        return df

    try:
        from rapidfuzz import fuzz, process
        have_rapidfuzz = True
    except ImportError:
        log.warning("rapidfuzz not installed; falling back to exact grouping.")
        have_rapidfuzz = False

    df = df.copy()
    name_source = df["payee"] if "payee" in df.columns else df["description"]
    df["clean_name"] = name_source.map(clean_name)
    df["group"] = df["clean_name"]

    if not have_rapidfuzz:
        return df

    for direction in ("paid", "received"):
        mask = df["direction"] == direction
        names = df.loc[mask, "clean_name"].tolist()
        canon = _cluster(names, fuzz, threshold)
        df.loc[mask, "group"] = [canon[n] for n in names]

    return df


def _cluster(names, fuzz, threshold):
    """Greedy clustering: each new name joins the first existing canonical name
    it is 'threshold' similar to, otherwise it starts its own cluster.
    Returns a dict {name -> canonical_name}."""
    canon_for = {}
    canonicals = []  # list of canonical strings, in first-seen order
    # Process most frequent names first so the canonical label is the common one.
    order = pd.Series(names).value_counts().index.tolist()
    for name in order:
        best_match, best_score = None, 0
        for c in canonicals:
            score = fuzz.token_sort_ratio(name, c)
            if score > best_score:
                best_match, best_score = c, score
        if best_match is not None and best_score >= threshold:
            canon_for[name] = best_match
        else:
            canonicals.append(name)
            canon_for[name] = name
    return canon_for


# --------------------------------------------------------------------------- #
# Ledger mapping (for the Tally export)
# --------------------------------------------------------------------------- #
def load_ledger_map(path: str) -> dict:
    """Load an editable CSV mapping of  group_name -> tally_ledger.
    Missing file -> empty mapping (everything uses the default ledger)."""
    if not path or not os.path.exists(path):
        return {}
    try:
        m = pd.read_csv(path, dtype=str, keep_default_na=False)
    except Exception as exc:
        log.warning("Could not read ledger map %s (%s)", path, exc)
        return {}
    cols = {c.strip().lower(): c for c in m.columns}
    gcol = cols.get("group_name") or cols.get("group") or list(m.columns)[0]
    lcol = cols.get("tally_ledger") or cols.get("ledger") or list(m.columns)[-1]
    return {
        str(r[gcol]).strip().upper(): str(r[lcol]).strip()
        for _, r in m.iterrows()
        if str(r[gcol]).strip()
    }


def attach_ledgers(df: pd.DataFrame, ledger_map: dict,
                   default_paid="Sundry Expenses",
                   default_received="Sundry Debtors") -> pd.DataFrame:
    """Add a 'ledger' column per row using the mapping, falling back to sensible
    defaults so nothing is ever dropped from the Tally export."""
    df = df.copy()

    def pick(row):
        key = str(row["group"]).strip().upper()
        if key in ledger_map and ledger_map[key]:
            return ledger_map[key]
        return default_paid if row["direction"] == "paid" else default_received

    df["ledger"] = df.apply(pick, axis=1)
    return df