"""parser.py - Data ingestion module for Hotel financial Excel files.
Parses:
  1. Detail Trial Balance (DTB)
  2. Income Statement Department (MTD)
  3. Consumption Report (Store issuing to Housekeeping & Guest Supplies)
  4. Laundry Reports (Bonvivo, Drop N Go)
"""

import logging
import os
import re
import glob
from typing import Dict, List, Any, Optional, Tuple
import openpyxl
import pandas as pd

# Configure module-level logger
logger = logging.getLogger(__name__)
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.INFO)


# Standard Account Code to Category mapping for Room Division / Housekeeping
ACCOUNT_MAP = {
    '0351010': 'Salaries & Wages',
    '0351012': 'Function Allowance',
    '0351016': 'Salaries & Wages (kary.lepas)',
    '0351040': 'Employess Transportation',
    '0351055': 'Service Charge (Account Perampungan PPh)',
    '0351060': 'Human Resources',
    '0351070': 'PTEB',
    '0352010': 'Uniforms',
    '0352020': 'Laundry & Dry Cleaning',
    '0352030': 'Printing & Stationery,Photo Copy, Postage & Stamp',
    '0352035': 'Meals',
    '0352040': 'Local Transport',
    '0352070': 'Music Cassette Entertainment & Tv Kabel',
    '0352100': 'Decoration',
    '0352200': 'Guest Supplies',
    '0352210': 'Cleaning Supplies',
    '0352220': 'Paper Suplies',
    '0352300': 'Linen Repl.',
    '0352310': 'China Glass & Silver Repl',
    '0352385': 'Outsourcing Utilities',
    '0352660': 'Telephone &  Internet',
    '0352980': 'Human Resources Other Expense',
}


# Month names for label inference: by leading month number and by name token
# (English + common Indonesian abbreviations).
_MONTH_NAMES_BY_NUMBER = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
]

_MONTH_NAMES_BY_TOKEN = {
    'jan': 'January', 'feb': 'February', 'mar': 'March', 'apr': 'April',
    'mei': 'May', 'may': 'May', 'jun': 'June', 'jul': 'July',
    'agu': 'August', 'ags': 'August', 'agt': 'August', 'aug': 'August',
    'sep': 'September', 'okt': 'October', 'oct': 'October',
    'nov': 'November', 'des': 'December', 'dec': 'December',
}


def infer_month_label(source: str, default: str = '') -> str:
    """
    Infers a canonical 'Month YYYY' label from a folder name, file name, or
    free-text label, normalizing English and Indonesian month names so that
    '7.JULY' -> 'July 2026' and '8.AGUSTUS' -> 'August 2026' everywhere.

    Resolution order:
      1. Leading month number  ('7.JULY' -> July, '08 AGUSTUS' -> August)
      2. Month-name token      ('agustus', 'aug', 'MEI', 'OKT', ...)
      3. `default` if provided, else the cleaned input text

    The year is taken from any 4-digit number in the input, else '2026'
    (the app's reporting period).
    """
    text = str(source or '').strip()
    if not text:
        return default

    year_match = re.search(r'(\d{4})', text)
    year = year_match.group(1) if year_match else '2026'
    without_year = re.sub(r'\d{4}', ' ', text)

    month_name = ''
    # An explicit month-name token wins over a leading number: in '09. Oktober'
    # the number may be a day or folder sequence, while 'Oktober' is unambiguous.
    for token in re.split(r'[^A-Za-z]+', without_year):
        name = _MONTH_NAMES_BY_TOKEN.get(token.strip().lower()[:3], '')
        if name:
            month_name = name
            break
    if not month_name:
        lead = re.match(r'^\s*(\d{1,2})\b', without_year)
        if lead and 1 <= int(lead.group(1)) <= 12:
            month_name = _MONTH_NAMES_BY_NUMBER[int(lead.group(1)) - 1]

    if month_name:
        return f"{month_name} {year}"
    return default if default else re.sub(r'\s+', ' ', without_year).strip()


def find_month_files(month_dir: str) -> Dict[str, Optional[str]]:
    """
    Scans a month folder (e.g., '8.AGUSTUS' or '7.JULY') and detects key Excel files.
    Returns a dict with file paths or None if not found.
    """
    files = {
        'dtb': None,
        'is_mtd': None,
        'consumption': None,
        'laundry_bonvivo': None,
        'laundry_dropngo': None,
        'bisrev': None,
    }

    if not month_dir or not os.path.isdir(month_dir):
        logger.warning(f"find_month_files: Directory not found or invalid: {month_dir}")
        return files

    logger.info(f"Scanning directory for Excel files: {month_dir}")
    search_paths = [month_dir]
    # Check subdirectories like 'Income Statement Dept HSD Agustus 2026' or 'Room Division Departement'
    for item in os.listdir(month_dir):
        sub = os.path.join(month_dir, item)
        if os.path.isdir(sub):
            search_paths.append(sub)

    for base in search_paths:
        for f in glob.glob(os.path.join(base, "*.xlsx")):
            fname = os.path.basename(f)
            if fname.startswith(".~lock"):
                continue
            lower = fname.lower()
            if "detail trial balance" in lower:
                files['dtb'] = f
                logger.debug(f"Found DTB file: {fname}")
            elif "income statement dept" in lower and "(mtd)" in lower:
                files['is_mtd'] = f
                logger.debug(f"Found Income Statement MTD file: {fname}")
            elif "income statement dept" in lower and files['is_mtd'] is None:
                files['is_mtd'] = f
                logger.debug(f"Found Income Statement file: {fname}")
            elif "consumption report" in lower:
                files['consumption'] = f
                logger.debug(f"Found Consumption Report file: {fname}")
            elif "bonvivo" in lower:
                files['laundry_bonvivo'] = f
                logger.debug(f"Found Bonvivo file: {fname}")
            elif "drop n go" in lower:
                files['laundry_dropngo'] = f
                logger.debug(f"Found Drop N Go file: {fname}")
            elif "bisrev" in lower and "pure" not in lower:
                files['bisrev'] = f
                logger.debug(f"Found Bisrev file: {fname}")

    # Log summary
    found_count = sum(1 for v in files.values() if v is not None)
    logger.info(f"find_month_files: Found {found_count} of 6 expected files in {month_dir}")

    return files


def parse_income_statement(is_path: str) -> pd.DataFrame:
    """
    Parses '08. Income Statement Dept - [Month] 2026 (MTD).xlsx'.
    Extracts Line Item, Category Group, Actual Amount (IDR), Budget Amount (IDR),
    Variance (IDR), Variance (%), and Actual Ratio (%).
    """
    logger.info(f"Parsing Income Statement: {is_path}")
    
    try:
        wb = openpyxl.load_workbook(is_path, data_only=True)
        target_sheets = [s for s in wb.sheetnames if 'ROOM' in s.upper() or 'LAUNDRY' in s.upper()]
        if not target_sheets:
            target_sheets = [wb.sheetnames[0]]
        records = []
    except Exception as e:
        logger.error(f"Error loading Income Statement workbook: {e}")
        raise

    for sname in target_sheets:
        ws = wb[sname]
        dept_name = "Room Division" if "ROOM" in sname.upper() else "Laundry Dept"
        current_group = "Other Expenses"

        for row_idx in range(1, ws.max_row + 1):
            col1 = ws.cell(row_idx, 1).value
            if not col1:
                continue
            
            line_str = str(col1).strip()
            
            # Detect group headers
            if "Payroll" in line_str:
                current_group = "Payroll & Related"
                continue
            elif "Other Expenses" in line_str:
                current_group = "Other Expenses"
                continue
            elif "Sales & Income" in line_str:
                current_group = "Revenue"
                continue
            elif "Cost of Good Sold" in line_str:
                current_group = "Cost of Goods Sold"
                continue
            
            # Detect subtotal / summary rows
            is_subtotal = any(k in line_str.lower() for k in [
                'total sales', 'total payroll', 'total other', 'total room',
                'total laundry', 'total cost', 'profit/loss'
            ])

            # Amounts are typically in Col 3 (Actual), Col 9 (Ratio), Col 11 (Budget)
            actual_val = ws.cell(row_idx, 3).value
            ratio_val = ws.cell(row_idx, 9).value
            budget_val = ws.cell(row_idx, 11).value

            # If not in standard columns, look across columns for numbers
            if actual_val is None or not isinstance(actual_val, (int, float)):
                candidates = []
                for c in range(2, 13):
                    v = ws.cell(row_idx, c).value
                    if isinstance(v, (int, float)) and v != 0:
                        candidates.append(v)
                if candidates:
                    actual_val = candidates[0]
                    budget_val = candidates[1] if len(candidates) > 1 else 0

            # Only process if we found numeric data
            if isinstance(actual_val, (int, float)):
                budget_num = float(budget_val) if isinstance(budget_val, (int, float)) else 0.0
                actual_num = float(actual_val)
                var_num = actual_num - budget_num
                var_pct = (var_num / budget_num * 100.0) if budget_num != 0 else 0.0
                ratio_num = float(ratio_val) if isinstance(ratio_val, (int, float)) else 0.0

                records.append({
                    'Department': dept_name,
                    'Group': current_group,
                    'Description': line_str,
                    'Actual': actual_num,
                    'Budget': budget_num,
                    'Variance': var_num,
                    'Variance_Pct': var_pct,
                    'Ratio_Pct': ratio_num,
                    'Is_Subtotal': is_subtotal,
                })

    return pd.DataFrame(records)


def parse_detail_trial_balance(dtb_path: str, filter_prefix: Optional[str] = '035') -> pd.DataFrame:
    """
    Parses '08. Detail Trial Balance HSD [Month] 2026.xlsx'.
    Extracts all individual transactions grouped by Account Code and Name.
    """
    logger.info(f"Parsing Detail Trial Balance: {dtb_path}")
    
    try:
        wb = openpyxl.load_workbook(dtb_path, data_only=True)
        ws = wb.active
        transactions = []
    except Exception as e:
        logger.error(f"Error loading Detail Trial Balance workbook: {e}")
        raise
        current_acct_code = None
        current_acct_name = None
        current_acct_full = None

        for row_idx, r in enumerate(ws.iter_rows(values_only=True), start=1):
            col0 = r[0]
            if not col0:
                continue
            
            str0 = str(col0).strip()
            
            # Check if row is an Account Header (starts with digit code of 5+ digits)
            acct_match = re.match(r'^(\d{5,})\s+(.*)', str0)
            if acct_match:
                code = acct_match.group(1)
                name = acct_match.group(2).strip()
                
                if filter_prefix and not code.startswith(filter_prefix):
                    current_acct_code = None
                    current_acct_name = None
                    current_acct_full = None
                    continue
                
                current_acct_code = code
                current_acct_name = name
                current_acct_full = str0
                continue

            # If we are inside an active account and row has a date in Col A
            if current_acct_code and re.match(r'^\d{4}-\d{2}-\d{2}', str0):
                # Columns in DTB:
                # 0: Date, 1: JRNL, 2: Partner, 3: Ref, 4: Source, 5: Dept, 6: Desc, 7: Beg, 8: Debit, 9: Credit, 10: Net, 11: End
                date_val = str0[:10]
                jrnl = str(r[1]) if r[1] is not None else ''
                partner = str(r[2]) if r[2] is not None else ''
                ref = str(r[3]) if r[3] is not None else ''
                source = str(r[4]) if r[4] is not None else ''
                dept = str(r[5]) if r[5] is not None else ''
                desc = str(r[6]) if r[6] is not None else ''
                debit = float(r[8]) if r[8] is not None and isinstance(r[8], (int, float)) else 0.0
                credit = float(r[9]) if r[9] is not None and isinstance(r[9], (int, float)) else 0.0
                net = debit - credit

                # Map to friendly category
                category = ACCOUNT_MAP.get(current_acct_code, current_acct_name)

                transactions.append({
                    'Date': date_val,
                    'Account_Code': current_acct_code,
                    'Account_Name': current_acct_name,
                    'Category': category,
                    'JRNL': jrnl,
                    'Partner': partner,
                    'Ref': ref,
                    'Source': source,
                    'Department': dept,
                    'Description': desc,
                    'Debit': debit,
                    'Credit': credit,
                    'Net_Amount': net,
                    'Row_Index': row_idx
                })

        df = pd.DataFrame(transactions)
        logger.info(f"parse_detail_trial_balance: Loaded {len(df)} transactions")
        return df
        
    except Exception as e:
        logger.error(f"Error parsing Detail Trial Balance: {e}")
        raise


def parse_consumption_report(cons_path: str) -> pd.DataFrame:
    """
    Parses '08. Consumption Report [Month] 2026.xlsx'.
    Reads both 'HK' and 'Guest Supplies' sheets to extract item-level transactions.
    """
    logger.info(f"Parsing Consumption Report: {cons_path}")
    
    try:
        wb = openpyxl.load_workbook(cons_path, data_only=True)
        items = []
    except Exception as e:
        logger.error(f"Error loading Consumption Report workbook: {e}")
        raise

    for sname in wb.sheetnames:
        ws = wb[sname]
        is_guest_supplies = "GUEST" in sname.upper()
        current_acct_section = "0352200 Guest Supplies" if is_guest_supplies else "General HK"

        for row in ws.iter_rows(values_only=True):
            if not row or not any(row):
                continue
            
            val0 = str(row[0]).strip() if row[0] is not None else ''
            
            # Detect section header in HK sheet (e.g. '03-52010 Uniforms')
            if '03-' in val0:
                current_acct_section = val0
                continue
            
            # Detect transaction line (starts with date YYYY-MM-DD)
            if re.match(r'^\d{4}-\d{2}-\d{2}', val0):
                # Standard format:
                # 0: Date, 1: No Trx, 2: Dept, 3: Process, 4: Type, 5: Item, 6: Qty, 7: Unit, 8: Price, 9: Amount
                trx_date = val0[:10]
                trx_no = str(row[1]).strip() if row[1] is not None else ''
                dept = str(row[2]).strip() if row[2] is not None else ''
                process = str(row[3]).strip() if row[3] is not None else ''
                trx_type = str(row[4]).strip() if row[4] is not None else ''
                item_name = str(row[5]).strip() if row[5] is not None else ''
                qty = float(row[6]) if row[6] is not None and isinstance(row[6], (int, float)) else 0.0
                unit = str(row[7]).strip() if row[7] is not None else ''
                price = float(row[8]) if row[8] is not None and isinstance(row[8], (int, float)) else 0.0
                amount = float(row[9]) if row[9] is not None and isinstance(row[9], (int, float)) else 0.0

                # Derive clean item title (strip prefix digits like 974059-Cleo -> Cleo)
                clean_name = re.sub(r'^\d+[\s\-_]+', '', item_name)

                items.append({
                    'Date': trx_date,
                    'Voucher_No': trx_no,
                    'Department': dept,
                    'Process': process,
                    'Type': trx_type,
                    'Item_Code_Name': item_name,
                    'Item_Name': clean_name,
                    'Qty': qty,
                    'Unit': unit,
                    'Unit_Price': price,
                    'Amount': amount,
                    'Sheet': sname,
                    'Section': current_acct_section
                })

    return pd.DataFrame(items)


def parse_laundry_reports(bonvivo_path: Optional[str], dropngo_path: Optional[str]) -> Dict[str, Any]:
    """
    Parses Bonvivo and Drop N Go laundry report Excel files.
    """
    logger.info(f"Parsing Laundry Reports - Bonvivo: {bonvivo_path}, Drop N Go: {dropngo_path}")
    
    result = {
        'bonvivo': {'total': 0.0, 'items': []},
        'dropngo': {'total': 0.0, 'items': []}
    }
    
    if bonvivo_path and os.path.exists(bonvivo_path):
        try:
            wb = openpyxl.load_workbook(bonvivo_path, data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                if "TOTAL" in sname.upper() or "GRAND" in sname.upper():
                    for r in ws.iter_rows(values_only=True):
                        for c in r:
                            if isinstance(c, (int, float)) and c > 100000:
                                result['bonvivo']['total'] = max(result['bonvivo']['total'], float(c))
            logger.info(f"Bonvivo total: {result['bonvivo']['total']}")
        except Exception as e:
            logger.error(f"Error parsing Bonvivo: {e}")

    if dropngo_path and os.path.exists(dropngo_path):
        try:
            wb = openpyxl.load_workbook(dropngo_path, data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                if "TOTAL" in sname.upper() or "GRAND" in sname.upper():
                    for r in ws.iter_rows(values_only=True):
                        for c in r:
                            if isinstance(c, (int, float)) and c > 100000:
                                result['dropngo']['total'] = max(result['dropngo']['total'], float(c))
            logger.info(f"Drop N Go total: {result['dropngo']['total']}")
        except Exception as e:
            logger.error(f"Error parsing Drop N Go: {e}")

    return result
