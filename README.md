⚖️ VendorSync — Vendor Reconciliation App
A Streamlit web application for Indian CAs and CFOs to reconcile Vendor Ledger vs Customer Ledger.
---
🚀 Quick Start
1. Install dependencies
```bash
pip install -r requirements.txt
```
2. Run the app
```bash
streamlit run app.py
```
The app will open at http://localhost:8501
---
📋 Input File Formats
Vendor Ledger (Excel)
Required columns (auto-detected):
Column	Description
Doc. Date	Document date
Doc No.	Document number
Doc Type Name	Type of document (Invoice, Debit Note, Payment, etc.)
Particulars	Description / narration
Opening	Opening balance
Debits	Debit amount
Credits	Credit amount
Closing Balance	Closing balance
Customer Ledger (Excel)
Required columns (auto-detected):
Column	Description
Document Date	Document date
Document Type	Type of document
Document no / Details	Document number / UTR / Cheque No
Debit (LC)	Debit amount in local currency
Credit (LC)	Credit amount in local currency
---
🔄 Matching Logic
#	Category	Matching Rules
1	Invoices	Matched by Document Number (exact, normalized)
2	Reversed Invoices	Detected (Debit = Credit in VL) → Excluded from matching
3	Debit Notes	1st: Doc Number · 2nd: Same Period + Same Amount (±tolerance)
4	Collections	1st: UTR Number · 2nd: Same Period + Same Amount (±tolerance)
5	Unmatched	All remaining items shown in Unmatched report
---
📊 Output
The app generates a downloadable Excel with these sheets:
Summary — Overview of all categories
Inv - Matched — Matched invoices
Inv - Unmatched VL / CL — Unmatched invoices per ledger
DN - Matched / Unmatched — Debit note results
Collections - Matched / Unmatched — Collection results
Reversed Excluded — Invoices excluded due to reversal
---
⚙️ Configuration
Amount Tolerance: Configurable in sidebar (default ₹1.00) — handles rounding differences
Column detection is automatic — no manual mapping needed
Debit Notes detected by Doc Type Name containing: `Debit Note`, `DN`, `Debit Memo`, `DM`
Collections detected by: `Payment`, `Receipt`, `NEFT`, `RTGS`, `IMPS`, `Cheque`, `TDS`, `UTR`
---
📦 Tech Stack
Streamlit — Web UI
Pandas — Data processing
OpenPyXL — Excel generation
NumPy — Numeric operations
