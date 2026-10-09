# Hotel Expense Tracker & SR Separator

A Streamlit application and automation engine designed for Hotel Room Division and Housekeeping teams. It reconciles monthly General Ledger accounting data with warehouse issuing logs and vendor reports, tracks item-level spending against departmental budgets, and separates warehouse Stock Request (SR) PDFs into operational supply categories.

---

## Overview

The application streamlines hotel financial reconciliation across two primary workflows:

1. **Monthly Expense Tracker**
   - Reconciles **Detail Trial Balance (DTB)**, **Income Statement (MTD/YTD)**, **Store Consumption Reports**, and **Laundry Reports** (Bonvivo, Drop N Go).
   - Automatically enriches General Ledger transactions with item details, quantities, unit prices, and store requisition voucher numbers.
   - Compares actual expenditures against departmental budget lines, identifying cost drivers and variance statuses (*On Track*, *Under Budget*, *Over Budget*, *Unbudgeted*).
   - Generates and exports comprehensive, multi-tab audit-ready Excel reports.

2. **Gudang Central SR Separator**
   - Ingests and parses warehouse Store Request (SR) PDF vouchers issued to Housekeeping.
   - Categorizes individual items into 4 operational supply categories:
     - 🧼 **Guest Supplies** (amenities, dental kits, bottled water, bath amenities, etc.)
     - 🧹 **Cleaning Supplies** (chemicals, trash bags, mops, sanitizers, gloves, etc.)
     - 🧻 **Paper Supplies** (facial tissues, toilet rolls, hand towels, laundry packaging bags, etc.)
     - 📄 **Print & Stationery** (guest comments, discrepancy slips, stationery, etc.)
   - Generates separated summary workbooks and downloadable Excel breakdowns.

---

## Architectural Highlights

- **Dynamic Sheet & Header Detection:** Intelligently identifies Room Division / Housekeeping sheets (`HK`, `Guest Supplies`, `FO`, `Welcome Drink`, etc.) across single or multi-department workbooks, automatically skipping non-relevant department sheets and pivot summaries.
- **Defensive Column & Row Access:** Safely parses varying tabular layouts without `IndexError`, accommodating variable column counts and truncated trailing cells.
- **Typo-Tolerant Reconciliation:** Uses normalization and token-based matching to align accounting template description quirks (e.g., matching template typo `"Paper Suplies"` with standard `"Paper Supplies"`).
- **Multi-Source Ingestion:** Seamlessly processes local monthly directory trees, user-uploaded Excel workbooks, or multi-month ZIP packages.
- **Layered Architecture:**
  - `app.py`: Application shell, auth gating, session management, and routing.
  - `views/`: View modules (`monthly_tracker.py`, `sr_separator.py`).
  - `parser.py`: Ingestion logic for DTB, Income Statement, Consumption Reports, and Laundry workbooks.
  - `matcher.py`: 100% transaction enrichment and budget reconciliation engine.
  - `exporter.py`: Excel workbook generation using OpenPyXL with stylized formatting.
  - `sr_parser.py`: Regex- and keyword-driven PDF extraction for warehouse store requisitions.
  - `data/`: Automatic month folder discovery and portable filesystem resolution.
  - `ui/`: Security guardrails, password protection, and Indonesian Rupiah currency formatting.

---

## Quick Start

### 1. Environment & Dependencies

Requires **Python 3.10+** (Python 3.12 recommended).

```bash
# Clone the repository and navigate into it
git clone https://github.com/cyborganeh/Hotel-Expense-Tracker.git
cd Hotel-Expense-Tracker

# Create and activate virtual environment
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
pip install pytest
```

### 2. Launch the Application

```bash
streamlit run app.py
```

Or using the convenience runner:

```bash
./run.sh
```

Open your browser at `http://localhost:8501`.

---

## CLI Usage: Standalone Expense Separation

For command-line automation, run `separate_expenses.py` to reconcile and generate the separated expense Excel workbook directly:

```bash
# Process a specific month folder
python separate_expenses.py "path/to/Business Review/8.AGUSTUS"

# Check script version
python separate_expenses.py --version
```

---

## Directory Setup & Environment Variables

The application automatically discovers month folders within default directories:
- `~/Documents/Business Review`
- `~/Business Review`
- Common Windows Documents paths

### Configurable Environment Variables:

| Variable | Description | Default |
|---|---|---|
| `HOTEL_DATA_DIR` | Custom root directory containing monthly folders | `~/Documents/Business Review` |
| `HOTEL_SR_DIR` | Custom root directory containing SR PDF files | `HOTEL_DATA_DIR/SR REPORT` |
| `HOTEL_APP_PASSWORD` | App login password override | Fallback: `123456` |

### Expected Folder Structure:

```text
Business Review/
├── 6. JUNE/
├── 7. JULY/
├── 8. AGUSTUS/
│   ├── 08. Detail Trial Balance HSD Agustus 2026.xlsx
│   ├── 08. Income Statement Dept - Agustus 2026 (MTD).xlsx
│   ├── 08. Consumption Report Agustus 2026.xlsx
│   ├── BONVIVO AGUSTUS 2026.xlsx
│   └── DROP N GO AGUSTUS 2026.xlsx
├── 9. SEP/
└── SR REPORT/
    ├── SR-001.pdf
    └── SR-002.pdf
```

---

## Authentication & Security

- **Password Gate:** Authenticates access via `ui/auth.py`. Credentials are resolved in the following priority:
  1. `.streamlit/secrets.toml` under `[auth] password = "..."`
  2. Environment variable `HOTEL_APP_PASSWORD`
  3. Default development password: `123456`
- **Path Sanitization:** File loaders reject traversal patterns, dangerous root directories, and unauthorized system access.
- **Upload Hardening:** Enforces strict MIME, size limits, and sanitization on uploaded files.

---

## Running Automated Tests

Run the complete test suite via `pytest`:

```bash
source .venv/bin/activate
pytest -v
```

The automated test suite verifies:
- Month folder and file discovery.
- Detail Trial Balance ingestion.
- Consumption report parsing with short rows, pivot tables, and multi-department workbooks.
- Budget vs. Actual reconciliation with category typo tolerances.
- Multi-tab Excel report generation and formatting.
- CLI argument handling and version flags.

---

## Project Structure

```text
amazing-mendel/
├── app.py                  # Main Streamlit entrance and navigation
├── app_launcher.py         # Entry point for packaged desktop deployment
├── parser.py               # Excel financial parser (DTB, IS, Consumption, Laundry)
├── matcher.py              # Reconciliation and transaction enrichment engine
├── exporter.py             # Formatted Excel export generation
├── sr_parser.py            # Warehouse Store Request PDF parser and classifier
├── separate_expenses.py    # Command-line interface for batch processing
├── build_windows_exe.py    # PyInstaller script for standalone Windows builds
├── run.sh                  # Convenience bash script to run the Streamlit app
├── requirements.txt        # Python dependency manifest
├── pyproject.toml          # Project metadata and test configuration
├── CHANGELOG.md            # Release changelog
├── README.md               # Documentation and usage guide
├── data/
│   └── months.py           # Month folder discovery, caching, and data combination
├── ui/
│   ├── auth.py             # Authentication gating
│   ├── formatting.py       # Currency (IDR) and numeric presentation formatters
│   └── security.py         # Path validation and file safety checks
├── views/
│   ├── monthly_tracker.py  # Monthly expense tracker & reconciliation UI view
│   └── sr_separator.py     # Warehouse SR PDF separator UI view
└── tests/
    └── test_core.py        # Core unit tests (discovery, parsing, matching, export)
```

---

## License

Internal proprietary software developed for Hotel Room Division & Housekeeping accounting management.
