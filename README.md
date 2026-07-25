# Bank Statement to Excel (+ Tally)

A simple Windows desktop tool. The user drags in **PDF / CSV / Excel** bank
statements, clicks **Convert**, and gets:

- **`Bank_Statement_*.xlsx`** — a grouped summary with drill-down:
  - **Summary** sheet — Column A/B = payments (name + total), Column C/D = receipts (name + total), sorted largest first. Each name links to its detail.
  - **Payments** / **Receipts** sheets — each payee is a collapsible **+/−** group; expand it to see every individual transaction with its date.
- **`Tally_Vouchers_*.xml`** (optional) — Payment/Receipt vouchers to import into Tally via *Gateway of Tally → Import → Vouchers*.

Files are saved to a **"Bank Statements"** folder on the Desktop.

---

## Project layout

```
bank-statement-tool/
├── main.py                 # the GUI (run this)
├── requirements.txt
├── config/
│   └── ledger_map.csv      # editable: payee name -> Tally ledger
└── core/
    ├── parsers.py          # PDF/CSV/Excel -> normalized transactions
    ├── normalize.py        # name cleaning + fuzzy grouping + ledger mapping
    ├── excel_report.py     # the collapsible Excel workbook
    ├── tally_export.py     # the Tally import XML
    └── pipeline.py         # orchestrates everything
```

## 1. Run it (development)

Install Python 3.11 or 3.12 (64-bit) from python.org, then in a terminal:

```bat
cd bank-statement-tool
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## 2. Build the Windows .exe

On the **Windows** machine, inside the activated venv:

```bat
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed ^
  --name "BankStatementTool" ^
  --collect-all pdfplumber ^
  --collect-all pdfminer ^
  --add-data "config;config" ^
  main.py
```

The exe appears in `dist\BankStatementTool.exe`. Ship that **together with the
`config` folder** next to it so users can edit `ledger_map.csv`.

`build.bat` in this folder does all of the above in one double-click.

> First launch may be slightly slow (PyInstaller unpacks itself). If Windows
> SmartScreen warns about an unknown publisher, that's normal for an unsigned
> exe — click *More info → Run anyway*, or code-sign it for wider distribution.

## 3. The Tally mapping (`config/ledger_map.csv`)

Two columns: `group_name,tally_ledger`. When a payee's grouped name matches
`group_name`, its voucher posts to that ledger. Anything unmatched falls back to
**Sundry Expenses** (payments) or **Sundry Debtors** (receipts).

**Every ledger you name here — plus the bank ledger — must already exist in
Tally**, or those vouchers fail on import. Grow this file over time as new
payees appear.

## Notes & tuning

- **PDF layouts vary by bank.** The parser tries table extraction first, then a
  line-by-line text fallback. If a particular bank extracts poorly, adjust
  `_pdf_table_settings` (line vs. text strategy) or the fallback regex in
  `core/parsers.py` — both are commented.
- **Fuzzy grouping strictness** is the `threshold` (default 88) in
  `normalize.group_transactions`. Lower = more aggressive merging.
- **Tally target is TallyPrime.** Tally ERP 9 uses an almost identical schema;
  if an import errors, verify the voucher tags against your exact version. The
  debit/credit sign convention is documented at the top of `tally_export.py`.
- **Scanned/image-only PDFs** have no text to extract and will be reported as
  skipped. Those need OCR (e.g. add an `ocrmypdf` pre-pass) — not included here.
