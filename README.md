# Hotel Expense Tracker & SR Separator

A Streamlit application for Room Division and Housekeeping teams. It reconciles monthly accounting data, tracks item-level spending, and separates warehouse stock request (SR) PDFs into operational supply categories.

---

## Overview

This project combines two workflows in one app:

- Monthly Expense Tracker: reconciles Detail Trial Balance (DTB), Income Statement, and Consumption data across month folders.
- SR Separator: parses Gudang Central SR PDFs and groups them into Guest Supplies, Cleaning Supplies, Paper Supplies, and Print & Stationery.

The app is intentionally organized into separate layers:

- `app.py` handles app shell, auth, theme, and route selection.
- `views/` contains dashboard-specific UI logic.
- `data/` contains month-folder discovery and default-path resolution.
- `ui/` contains auth, formatting, and security helpers.

---

## Key Features

1. Monthly reconciliation dashboard
   - Matches DTB transactions to departmental budget lines.
   - Combines actuals from multiple sources into a single monthly view.
   - Renders executive KPIs and category variance summaries.

2. Multi-source data loading
   - Local month folders under a chosen Business Review directory.
   - Uploaded month Excel files.
   - ZIP archives containing multiple month folders.

3. Excel export pipeline
   - Generates multi-tab separated workbooks for monthly reporting.
   - Supports download of category-specific and combined outputs.

4. SR PDF processing
   - Reads SR PDFs from a local folder or uploaded files.
   - Summarizes totals, item cost drivers, and category breakdowns.
   - Exports a separated SR workbook.

5. Authentication and hardened file handling
   - Password-protected app access through the auth gate.
   - Directory validation blocks dangerous system paths.
   - Upload validation prevents oversized or unsafe files.

---

## Quick Start

### 1. Install dependencies

This project requires Python 3.10 or newer.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m pip install pytest
```

### 2. Run the app

```bash
streamlit run app.py
```

Or use the convenience script:

```bash
./run.sh
```

Then open the local Streamlit URL, usually:

```text
http://localhost:8501
```

---

## Data Directory Setup

The app looks for month data under a parent folder and automatically checks common locations such as:

- `~/Documents/Business Review`
- `~/Business Review`

You can override this with the environment variable:

```bash
export HOTEL_DATA_DIR="/path/to/Business Review"
```

If the value is invalid or points to a blocked system path, the app ignores it and falls back to the normal default search instead of crashing.

Expected folder layout:

```text
Business Review/
├── 8.AGUSTUS/
│   ├── Detail Trial Balance.xlsx
│   ├── Income Statement MTD.xlsx
│   └── Consumption Report.xlsx
├── 7.JULY/
├── 6.JUNE/
└── SR REPORT/
    ├── SR-001.pdf
    └── SR-002.pdf
```

Notes:
- Each month folder is typically discovered automatically.
- The app warns when a selected directory is missing or invalid.
- SR files can also be loaded from `HOTEL_SR_DIR` if that environment variable is set.

---

## Authentication

The app includes a simple password gate in `ui/auth.py`.

Password resolution order:

1. `.streamlit/secrets.toml` under `[auth] password = "..."`
2. `HOTEL_APP_PASSWORD`
3. built-in fallback password: `123456`

This is intended for local use and should be changed before sharing the app.

---

## Running Tests

```bash
source .venv/bin/activate
python -m pytest tests/test_core.py -q
```

The current test suite covers month discovery, parsing expectations, reconciliation flow, and Excel export behavior.

---

## Project Structure

```text
amazing-mendel/
├── app.py
├── parser.py
├── matcher.py
├── exporter.py
├── sr_parser.py
├── ui_components.py
├── separate_expenses.py
├── run.sh
├── requirements.txt
├── pyproject.toml
├── README.md
├── data/
│   └── months.py
├── ui/
│   ├── auth.py
│   ├── formatting.py
│   └── security.py
├── views/
│   ├── monthly_tracker.py
│   └── sr_separator.py
├── tests/
│   └── test_core.py
└── data/
    └── months.py
```

---

## Notes

- Generated Excel files are stored in the output path you provide or downloaded from the app UI.
- The app uses Indonesian Rupiah formatting and hotel-specific category matching rules.
- The startup path now gracefully falls back when `HOTEL_DATA_DIR` is invalid instead of crashing the app.
