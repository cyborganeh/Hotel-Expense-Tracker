# Hotel Santika Depok - Expense Separator & Spending Tracker

A specialized financial application designed for Hotel Santika Depok (Room Division & Housekeeping) to automatically separate, reconcile, and track monthly expenses from complex accounting Excel files.

---

## 🌟 Key Features

1. **100% Automated Reconciliation**:
   - Matches General Ledger accounts (`0351010`–`0352980`) directly with official departmental Income Statement budget lines.
   - Links store issuing vouchers (`GDC/OUT/...`) to individual product lines in the Consumption Report (e.g. Cleo water, toilet paper, slippers, chemicals).
   - Extracts and links vendor contracts (PT Pandu Jasa Terpadu cleaning service, Bonvivo laundry, Drop N Go, Maxindo Internet, Indovision TV cable).

2. **One-Click Multi-Tab Separated Excel Exporter**:
   - Generates a styled workbook with currency formatting (`Rp #,##0`), auto-width columns, and color-coded variance tags:
     - 📑 **Executive Summary**: KPI metrics, budget vs. actual variance table, status badges.
     - 📑 **All Transactions**: Master audit journal with 100% of classified transactions.
     - 📑 **Guest Supplies**: Itemized issuing for Cleo water, slippers, toothbrush, soap, etc.
     - 📑 **Outsourcing & Laundry**: Cleaning service, Bonvivo, Drop N Go.
     - 📑 **Cleaning Supplies**: Cleaning chemicals, garbage bags, air freshener contract.
     - 📑 **Paper Supplies**: Toilet rolls, facial tissue, laundry packaging bags.
     - 📑 **Payroll & SC**: Basic salary, daily workers, transportation allowance, service charge.
     - 📑 **Media & Utilities**: Internet (Maxindo), TV cable (MNC Sky Vision), Telkom.

3. **Interactive Web Dashboard (Streamlit & Plotly)**:
   - **Executive Dashboard**: Donut chart of category distribution, top cost drivers, budget vs actual comparison.
   - **Transaction Drilldown**: Search by product name, vendor, or voucher number with instant subtotal calculation.
   - **Budget vs. Actual Variance**: Highlights over-budget categories (e.g. Music/TV +113.9%).
   - **Month-over-Month Comparison**: Compares August 2026 vs July 2026 trends.
   - **Data Exporter**: Download the complete separated workbook or individual category slices (CSV/Excel).

---

## 🚀 Quick Start

### 1. Launch Interactive Web App
```bash
./run.sh
```
Or directly with uv:
```bash
uv run streamlit run app.py
```
Open **http://localhost:8501** in your web browser.

### 2. Run Headless CLI Separator
To separate a specific month folder directly into an Excel workbook:
```bash
uv run python3 separate_expenses.py --input "/home/rzl/Documents/Business Review/8.AGUSTUS" --output "Separated_Expenses_Agustus_2026.xlsx"
```

To process all available months (August, July, June) in one command:
```bash
uv run python3 separate_expenses.py --all
```

---

## 📁 Project Structure

```
amazing-mendel/
├── app.py                     # Streamlit web application
├── parser.py                  # Financial Excel parser (DTB, IS, Consumption, Laundry)
├── matcher.py                 # Reconciliation & transaction enrichment engine
├── exporter.py                # Multi-tab openpyxl Excel generator
├── separate_expenses.py       # Standalone CLI batch runner
├── run.sh                     # One-click startup shell script
├── pyproject.toml             # Dependencies configuration (uv / pip)
├── Separated_Expenses_Agustus_2026.xlsx # Generated August workbook
├── Separated_Expenses_Juli_2026.xlsx    # Generated July workbook
└── README.md                  # Documentation
```
