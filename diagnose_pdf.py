"""
diagnose_pdf.py
---------------
Run this ONLY if a bank PDF still isn't reading correctly. It prints exactly
what the tool "sees" inside the PDF so the layout can be pinned precisely.

USAGE (from inside the project folder, with the venv active):

    python diagnose_pdf.py "C:\\path\\to\\your\\statement.pdf"

Then copy the whole printed output and share it. Replace any real names or
account numbers with fake ones first if you like — only the column POSITIONS
matter for fixing alignment.
"""

import sys
from core import parsers


def main():
    if len(sys.argv) < 2:
        print("Usage: python diagnose_pdf.py <path-to-pdf>")
        return
    path = sys.argv[1]

    try:
        import pdfplumber
    except ImportError:
        print("pdfplumber is not installed. Run: pip install pdfplumber")
        return

    with pdfplumber.open(path) as pdf:
        print(f"PAGES: {len(pdf.pages)}")
        page = pdf.pages[0]
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        lines = parsers._group_words_into_lines(words)

        anchors = parsers._find_amount_anchors(lines)
        print(f"\nDETECTED AMOUNT-COLUMN ANCHORS: {anchors}")
        print("(If this is None, the Withdrawal/Deposit/Balance header wasn't "
              "found — that's the problem to fix.)")

        print("\nFIRST 20 LINES (each word shown as text@x0-x1):")
        for i, ln in enumerate(lines[:20]):
            parts = [f"{w['text']}@{int(w['x0'])}-{int(w['x1'])}" for w in ln]
            print(f"  L{i:02d} top={int(ln[0]['top']):4d}: " + " | ".join(parts))

    print("\nPARSED RESULT (what the tool extracts):")
    df = parsers.parse_file(path)
    if df.empty:
        print("  (no transactions extracted)")
    else:
        cols = [c for c in ["date", "payee", "amount", "direction", "balance"]
                if c in df.columns]
        print(df[cols].head(15).to_string(index=False))
        print(f"\n  Total transactions extracted: {len(df)}")


if __name__ == "__main__":
    main()
