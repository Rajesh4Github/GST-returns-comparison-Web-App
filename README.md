# GST-returns-comparison-Web-App
📘 Introduction
This application is a GST audit and reconciliation app that automates the extraction, cleaning, and consolidation of multiple GST return and ledger statements. It is designed to help taxpayers and auditors reconcile outward supplies, input tax credit (ITC), and tax liabilities across different statutory returns and ledgers.

🎯 Coverage
The workflow integrates data from:

GSTR‑1 → Outward supplies (sales) reported.

GSTR‑2B → Input tax credit available and not available.

GSTR‑3B → Monthly summary return of outward supplies, ITC claimed, and tax liability.

Electronic Credit Ledger → Tracks ITC balances available for utilization.

Electronic Cash Ledger → Records cash deposits and payments made towards GST liability.

Electronic Credit Reversal & Re‑claimed Statement → Captures ITC reversals and subsequent re‑claims.

RCM Liability/ITC Statement → Details tax payable under reverse charge mechanism and corresponding ITC eligibility.

⚙️ Objectives
Automated PDF Parsing → Extract structured data from GST returns and ledger PDFs using pdfplumber and pypdf.

Data Cleaning & Standardization → Convert raw values into consistent numeric formats.

Month Mapping → Identify the tax period (APR–MAR) from filenames or PDF text.

Template Alignment → Populate predefined layouts for each return/ledger type.

Consolidation → Merge GSTR‑1, GSTR‑2B, GSTR‑3B, and ledger statements into a single Excel workbook with separate sheets.

Reconciliation → Provide a unified view to compare outward supplies, ITC availability, reversals, and liabilities across returns and ledgers.

🛠️ Tools & Libraries
pdfplumber / pypdf → PDF text and table extraction

pandas → DataFrame manipulation and reconciliation

openpyxl → Export consolidated outputs into Excel

regex (re) → Clean and parse numeric values