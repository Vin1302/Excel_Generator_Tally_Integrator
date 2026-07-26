"""
tally_export.py
---------------
Generate a Tally import XML file containing Payment and Receipt vouchers, one per
transaction, ready for  Gateway of Tally -> Import -> Vouchers.

Sign convention used (standard Tally):
    AMOUNT negative  = Debit  (Dr)   , ISDEEMEDPOSITIVE = Yes
    AMOUNT positive  = Credit (Cr)   , ISDEEMEDPOSITIVE = No

Payment voucher (money OUT of the bank):
    party/expense ledger : Debit  (Dr)
    bank ledger          : Credit (Cr)

Receipt voucher (money INTO the bank):
    bank ledger          : Debit  (Dr)
    party ledger         : Credit (Cr)

IMPORTANT
  * Every ledger named here (bank ledger + each party/expense ledger) MUST
    already exist in the target Tally company, or those vouchers fail on import.
    The ledger names come from normalize.attach_ledgers().
  * This targets TallyPrime. Tally ERP 9 uses an almost identical structure; if
    you hit an import error, verify the tags against your exact version before
    assuming the data is wrong.
"""

from __future__ import annotations

from xml.sax.saxutils import escape
import pandas as pd


def _amt(value: float, debit: bool) -> str:
    v = abs(float(value))
    return f"-{v:.2f}" if debit else f"{v:.2f}"


def _ledger_entry(ledger_name: str, amount: float, debit: bool) -> str:
    return f"""        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{escape(ledger_name)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>{"Yes" if debit else "No"}</ISDEEMEDPOSITIVE>
          <AMOUNT>{_amt(amount, debit)}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""


def _voucher(row, bank_ledger: str) -> str:
    date = pd.to_datetime(row["date"]).strftime("%Y%m%d")
    amount = float(row["amount"])
    party = str(row["ledger"])
    narration = escape(str(row.get("description", ""))[:250])

    if row["direction"] == "paid":
        vtype = "Payment"
        # party Dr, bank Cr
        entries = (_ledger_entry(party, amount, debit=True)
                   + "\n"
                   + _ledger_entry(bank_ledger, amount, debit=False))
    else:
        vtype = "Receipt"
        # bank Dr, party Cr
        entries = (_ledger_entry(bank_ledger, amount, debit=True)
                   + "\n"
                   + _ledger_entry(party, amount, debit=False))

    deemed = "Yes" if row["direction"] == "paid" else "No"
    voucher = f"""      <VOUCHER VCHTYPE="{vtype}" ACTION="Create" OBJVIEW="Accounting Voucher View">
        <DATE>{date}</DATE>
        <EFFECTIVEDATE>{date}</EFFECTIVEDATE>
        <VOUCHERTYPENAME>{vtype}</VOUCHERTYPENAME>
        <NARRATION>{narration}</NARRATION>
        <ISDEEMEDPOSITIVE>{deemed}</ISDEEMEDPOSITIVE>
{entries}
      </VOUCHER>"""
    # Each voucher must sit inside its own <TALLYMESSAGE>.
    return ('    <TALLYMESSAGE xmlns:UDF="TallyUDF">\n'
            + voucher + "\n    </TALLYMESSAGE>")


def build_tally_xml(df: pd.DataFrame, out_path: str,
                    company_name: str, bank_ledger: str) -> str:
    """df must have: date, description, amount, direction, ledger."""
    messages = "\n".join(_voucher(r, bank_ledger) for _, r in df.iterrows())

    xml = f"""<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>{escape(company_name)}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
{messages}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>"""

    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(xml)
    return out_path
