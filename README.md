# Hotel - Expense Separator & Spending Tracker

A Streamlit application for Room Division and Housekeeping teams. It reconciles monthly accounting workbooks, tracks item-level spending, and separates Warehouse Central Stock Request (SR) PDFs into operational supply categories.

---

## 🌟 Key Features

1. **Automated Monthly Expense Reconciliation**:
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

3. **Interactive Monthly Expense Dashboard (Streamlit & Plotly)**:
   - **Executive Dashboard**: Donut chart of category distribution, top cost drivers, budget vs actual comparison.
   - **Transaction Drilldown**: Search by product name, vendor, or voucher number with instant subtotal calculation.
   - **Budget vs. Actual Variance**: Highlights over-budget categories (e.g. Music/TV +113.9%).
   - **Month-over-Month Comparison**: Compares August 2026 vs July 2026 trends.
   - **Data Exporter**: Download the complete separated workbook or individual category slices (CSV/Excel).

4. **Warehouse Stock Request (SR) Separator**:
   - Loads SR PDFs from the local `SR REPORT` folder or accepts multiple uploaded PDFs.
   - Categorizes issued items into Guest Supplies, Cleaning Supplies, Paper Supplies, and Print & Stationery.
   - Shows category totals, repeated item orders, issuing history, cost drivers, and filtered logs.
   - Exports a separated multi-tab SR workbook.

5. **Flexible Data Sources**:
   - Select month folders from a local Business Review directory.
   - Upload one month at a time or upload a ZIP containing multiple month folders.

---

## 🚀 Quick Start

### 1. Install dependencies
The project requires Python 3.10 or newer. With [uv](https://docs.astral.sh/uv/) installed, dependencies are created and installed automatically when the app starts. To install them explicitly:

```bash
uv sync
```

Alternatively, install the runtime dependencies with pip:

```bash
python -m pip install -r requirements.txt
```

### 2. Launch Interactive Web App
```bash
./run.sh
```
Or directly with uv:
```bash
uv run streamlit run app.py
```
Open **http://localhost:8501** in your web browser.

In the sidebar, choose **Monthly Expense Tracker** or **Stock Request (SR) Separator**. The tracker can use local month folders, uploaded Excel files, or uploaded ZIP archives. The SR separator accepts PDF uploads or local PDFs under `SR REPORT`.

### 3. Run Headless CLI Separator
To separate a specific month folder directly into an Excel workbook:
```bash
uv run python3 separate_expenses.py --input "/home/rzl/Documents/Business Review/8.AGUSTUS" --output "Separated_Expenses_Agustus_2026.xlsx"
```

To process all available months (August, July, June) in one command:
```bash
uv run python3 separate_expenses.py --all
```

Show the installed CLI version:
```bash
uv run python3 separate_expenses.py --version
```

The CLI also accepts an optional month label:

```bash
uv run python3 separate_expenses.py --input "/path/to/month-folder" --month "August 2026"
```

### 4. Run automated checks
```bash
uv run --extra dev python -m pytest tests/ -v
```

## Data Folder Layout

By default, the application looks for a Business Review folder in `~/Documents/Business Review` or `~/Business Review`. Set `HOTEL_DATA_DIR` to use another parent directory. Each month folder should contain the relevant Excel workbooks, including a Detail Trial Balance file; Income Statement and Consumption Report files are optional.

For local Stock Request PDFs, place files in `SR REPORT` under the selected data directory, or set `HOTEL_SR_DIR` to a different folder:

```text
Business Review/
├── 8.AGUSTUS/
│   ├── Detail Trial Balance.xlsx
│   ├── Income Statement MTD.xlsx
│   └── Consumption Report.xlsx
├── 7.JULY/
└── SR REPORT/
   ├── SR-001.pdf
   └── SR-002.pdf
```

Uploaded files are processed in memory or temporary files and are not added to the repository.

---

## 📁 Project Structure

```
amazing-mendel/
├── app.py                     # Streamlit web application
├── parser.py                  # Financial Excel parser (DTB, IS, Consumption, Laundry)
├── matcher.py                 # Reconciliation & transaction enrichment engine
├── exporter.py                # Multi-tab openpyxl Excel generator
├── sr_parser.py               # Stock Request PDF parser and summaries
├── ui_components.py            # Shared dashboard charts and KPI components
├── requirements.txt            # pip runtime dependencies
├── separate_expenses.py       # Standalone CLI batch runner
├── run.sh                     # One-click startup shell script
├── pyproject.toml             # Dependencies configuration (uv / pip)
├── tests/
│   └── test_core.py            # Automated tests
└── README.md                  # Documentation
```

## Notes

- Generated Excel files are written to the path supplied to the CLI or downloaded from the web app; they are not required to run the application.
- The application uses Indonesian Rupiah (IDR) formatting and hotel-specific category matching rules.
