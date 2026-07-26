"""
pipeline.py
-----------
Ties everything together:

    files  ->  parse  ->  group (fuzzy)  ->  Excel report
                                        ->  (optional) Tally XML

Call run() with a progress callback and it returns a small result object the GUI
can display.
"""

from __future__ import annotations

import os
import datetime as dt
from dataclasses import dataclass, field
from typing import Callable, List, Optional

from . import parsers, normalize, excel_report, tally_export


@dataclass
class Result:
    excel_path: Optional[str] = None
    tally_path: Optional[str] = None
    n_transactions: int = 0
    n_paid: int = 0
    n_received: int = 0
    skipped_files: List[str] = field(default_factory=list)
    message: str = ""


def run(
    files: List[str],
    output_dir: str,
    make_tally: bool = False,
    company_name: str = "",
    bank_ledger: str = "Bank",
    ledger_map_path: str = "",
    progress: Callable[[int, str], None] = lambda pct, msg: None,
) -> Result:
    os.makedirs(output_dir, exist_ok=True)
    res = Result()

    # 1. Parse every file --------------------------------------------------- #
    progress(5, "Reading your files...")
    frames = []
    for i, path in enumerate(files):
        df = parsers.parse_file(path)
        if df is None or df.empty:
            res.skipped_files.append(os.path.basename(path))
        else:
            frames.append(df)
        progress(5 + int(35 * (i + 1) / max(len(files), 1)),
                 f"Reading {os.path.basename(path)}")

    if not frames:
        res.message = ("Couldn't read any transactions. The files may be scanned "
                       "images or use an unusual layout.")
        return res

    import pandas as pd
    tx = pd.concat(frames, ignore_index=True).sort_values("date")

    # 2. Group similar payees ---------------------------------------------- #
    progress(55, "Grouping similar payments...")
    tx = normalize.group_transactions(tx)

    res.n_transactions = len(tx)
    res.n_paid = int((tx["direction"] == "paid").sum())
    res.n_received = int((tx["direction"] == "received").sum())

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    # 3. Excel report ------------------------------------------------------- #
    progress(70, "Building the Excel report...")
    excel_path = os.path.join(output_dir, f"Bank_Statement_{stamp}.xlsx")
    excel_report.build_workbook(tx, excel_path)
    res.excel_path = excel_path

    # 4. Tally XML (optional) ---------------------------------------------- #
    if make_tally:
        progress(88, "Creating the Tally import file...")
        ledger_map = normalize.load_ledger_map(ledger_map_path)
        tx_led = normalize.attach_ledgers(tx, ledger_map)
        tally_path = os.path.join(output_dir, f"Tally_Vouchers_{stamp}.xml")
        tally_export.build_tally_xml(
            tx_led, tally_path,
            company_name=company_name or "My Company",
            bank_ledger=bank_ledger or "Bank",
        )
        res.tally_path = tally_path

    progress(100, "Done")
    parts = [f"{res.n_transactions} transactions "
             f"({res.n_paid} paid, {res.n_received} received)."]
    if res.skipped_files:
        parts.append("Could not read: " + ", ".join(res.skipped_files))
    res.message = " ".join(parts)
    return res
