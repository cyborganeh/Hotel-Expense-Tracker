"""
sr_parser.py - Stock Request (SR) PDF Parser and Categorizer
Hotel - Room Division & Housekeeping

Parses Stock Request Consumption PDF files issued by Gudang Central (Warehouse)
and automatically categorizes items into:
  1. Guest Supplies
  2. Cleaning Supplies
  3. Paper Supplies
  4. Print & Stationery
"""

import os
import re
import io
from typing import Dict, List, Any, Optional, Union
import pandas as pd
import pymupdf


# Complete Item Code to Category Mapping based on Hotel SR tracking
ITEM_CATEGORY_MAP = {
    # --- GUEST SUPPLIES ---
    '976729': 'Guest Supplies',  # Cotton Bud
    '976724': 'Guest Supplies',  # Vanity Kit
    '976731': 'Guest Supplies',  # Shower Cap
    '976732': 'Guest Supplies',  # Sanitary Bag
    '976717': 'Guest Supplies',  # Bath Shower Gel Dirigen
    '976718': 'Guest Supplies',  # Conditioning Shampo Dirigen
    '976744': 'Guest Supplies',  # Soap 25 Gr
    '976747': 'Guest Supplies',  # Toothbrush
    '976728': 'Guest Supplies',  # Toothbrush Bamboo
    '976743': 'Guest Supplies',  # Slipper Logo
    '976737': 'Guest Supplies',  # Shoe Shine
    '976592': 'Guest Supplies',  # Coaster
    '976781': 'Guest Supplies',  # Memo Pad 1/3 Folio
    '976800': 'Guest Supplies',  # Plastik Laundry
    '976715': 'Guest Supplies',  # Laundry Bag Logo
    '976620': 'Guest Supplies',  # White Sugar Sachet
    '976597': 'Guest Supplies',  # Brown Sugar Logo
    '976741': 'Guest Supplies',  # Coffee Logo
    '976742': 'Guest Supplies',  # Tea Logo
    '976600': 'Guest Supplies',  # Creamer Logo
    '976589': 'Guest Supplies',  # Stirer
    '976721': 'Guest Supplies',  # Water Tag Cleo 330Ml
    '974059': 'Guest Supplies',  # Cleo 330 Ml
    '976608': 'Guest Supplies',  # Botol Cleo 330Ml
    '976780': 'Guest Supplies',  # Guest Comment Room
    '977197': 'Guest Supplies',  # Green Tag Bathroom
    '977198': 'Guest Supplies',  # Green Tag Bed

    # --- CLEANING SUPPLIES ---
    '981641': 'Cleaning Supplies',  # Plastik Sampah Hitam 90 x 120

    # --- PAPER SUPPLIES ---
    '976802': 'Paper Supplies',  # Facial Tissue - Special White
    '976799': 'Paper Supplies',  # Hand Towel Tissue
    '976801': 'Paper Supplies',  # Toilet Tissue 85Gr@Roll
    '976811': 'Paper Supplies',  # Plastik Roll

    # --- PRINT & STATIONERY ---
    '976933': 'Print & Stationery',  # Isi Stapler No.10
}

# Fallback Keyword to Category Mapping for any unlisted items
KEYWORD_CATEGORY_MAP = [
    # Paper Supplies
    (re.compile(r'tissue|toilet\s+roll|hand\s+towel|facial\s+tissue|plastik\s+roll', re.I), 'Paper Supplies'),
    # Cleaning Supplies
    (re.compile(r'plastik\s+sampah|glass\s+cleaner|all\s+purpose|metal\s+shine|scale\s+off|carpet\s+shampoo|smooth|air\s+frezz|bathklin|disinfectant|chemical|detergent|mop|broom', re.I), 'Cleaning Supplies'),
    # Print & Stationery
    (re.compile(r'stapler|bill\s+|kertas|hvs|lakban|spidol|tinta|buku\s+folio|stabillo|isolasi|pen\b|paper\s+clip|amplop|form\b', re.I), 'Print & Stationery'),
    # Guest Supplies (amenities, beverages, hotel linen tags, guest bags)
    (re.compile(r'cleo|water|sugar|coffee|tea|creamer|stirer|soap|shampoo|shower|slipper|toothbrush|dental|comb|razor|shaving|vanity|cotton\s+bud|sanitary|shoe\s+shine|coaster|memo|laundry\s+bag|green\s+tag|tag\b', re.I), 'Guest Supplies'),
]


def clean_number(val: Any) -> float:
    """Converts strings like '1,440.00' or '269,215' or floats to float."""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(',', '').replace(' ', '')
    try:
        return float(s)
    except ValueError:
        return 0.0


def categorize_item(item_code: str, item_name: str) -> str:
    """
    Determines category for an item:
    1. Exact item code lookup in ITEM_CATEGORY_MAP
    2. Keyword regex rules on item_name
    3. Default to 'Guest Supplies' (Housekeeping dominant)
    """
    code_str = str(item_code).strip()
    if code_str in ITEM_CATEGORY_MAP:
        return ITEM_CATEGORY_MAP[code_str]

    name_lower = str(item_name).strip().lower()
    for pattern, cat in KEYWORD_CATEGORY_MAP:
        if pattern.search(name_lower):
            return cat

    return 'Guest Supplies'


def parse_sr_pdf(source: Union[str, bytes, Any]) -> pd.DataFrame:
    """
    Parses a single Stock Request Consumption PDF.
    Supports file path, bytes, or file-like object (Streamlit UploadedFile).
    Returns DataFrame of line items.
    """
    if isinstance(source, (str, os.PathLike)):
        doc = pymupdf.open(source)
        source_name = os.path.basename(source)
    elif hasattr(source, "read"):
        content = source.read()
        if hasattr(source, "seek"):
            source.seek(0)
        doc = pymupdf.open(stream=content, filetype="pdf")
        source_name = getattr(source, "name", "Uploaded_SR.pdf")
    elif isinstance(source, (bytes, bytearray)):
        doc = pymupdf.open(stream=source, filetype="pdf")
        source_name = "Uploaded_SR.pdf"
    else:
        raise ValueError("Unsupported source type for parse_sr_pdf")

    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    # Extract metadata
    sr_match = re.search(r'Stock Request Consumption\s+([A-Z0-9/]+)', full_text)
    if not sr_match:
        sr_match = re.search(r'Request Reference:\s*\n\s*([A-Z0-9/]+)', full_text)
    sr_number = sr_match.group(1).strip() if sr_match else source_name.replace('.pdf', '')

    date_match = re.search(r'Creation Date:\s*\n?\s*([0-9/]{8,10})', full_text)
    if date_match:
        raw_date = date_match.group(1).strip()
        parts = raw_date.split('/')
        if len(parts) == 3:
            date_str = f"{parts[2]}-{parts[0].zfill(2)}-{parts[1].zfill(2)}"
        else:
            date_str = raw_date
    else:
        date_str = "Unknown"

    req_match = re.search(r'Requested by:\s*\n?\s*([^\n]+)', full_text)
    requested_by = req_match.group(1).strip() if req_match else "Housekeeping"

    dept_match = re.search(r'Department:\s*\n?\s*([^\n]+)', full_text)
    department = dept_match.group(1).strip() if dept_match else "03 - House Keeping"

    # Parse product rows
    lines = [l.strip() for l in full_text.split('\n') if l.strip()]
    records = []

    for idx, line in enumerate(lines):
        # Match lines like '976802 - Facial Tissue - Special White'
        m = re.match(r'^(\d{5,8})\s*-\s*(.+)$', line)
        if m:
            code = m.group(1).strip()
            name = m.group(2).strip()

            qty = 1.0
            unit = "Pcs"
            cost = 0.0
            total = 0.0

            try:
                qty = clean_number(lines[idx + 1])
                unit = lines[idx + 2]
                cost = clean_number(lines[idx + 3])
                total = clean_number(lines[idx + 4])
            except (IndexError, ValueError):
                pass

            category = categorize_item(code, name)

            records.append({
                'SR_Number': sr_number,
                'Date': date_str,
                'Requested_By': requested_by,
                'Department': department,
                'Item_Code': code,
                'Item_Name': name,
                'Category': category,
                'Qty': qty,
                'Unit': unit,
                'Cost': cost,
                'Total': total,
                'Source_File': source_name
            })

    return pd.DataFrame(records)


def parse_multiple_sr_pdfs(sources: List[Union[str, bytes, Any]]) -> pd.DataFrame:
    """
    Parses multiple SR PDFs and returns a combined sorted DataFrame.
    """
    dfs = []
    for s in sources:
        try:
            df = parse_sr_pdf(s)
            if not df.empty:
                dfs.append(df)
        except Exception as e:
            print(f"Error parsing SR PDF: {e}")

    if not dfs:
        return pd.DataFrame(columns=[
            'SR_Number', 'Date', 'Requested_By', 'Department',
            'Item_Code', 'Item_Name', 'Category', 'Qty', 'Unit', 'Cost', 'Total', 'Source_File'
        ])

    combined_df = pd.concat(dfs, ignore_index=True)
    combined_df = combined_df.sort_values(by=['Date', 'SR_Number', 'Category', 'Item_Name']).reset_index(drop=True)
    return combined_df


def get_category_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns summary statistics aggregated by category.
    """
    if df.empty:
        return pd.DataFrame()

    total_amount_all = df['Total'].sum()

    cat_sum = df.groupby('Category').agg(
        Unique_Items=('Item_Name', 'nunique'),
        Total_Qty=('Qty', 'sum'),
        Total_Amount=('Total', 'sum'),
        Order_Count=('Total', 'count')
    ).reset_index()

    cat_sum['Pct_Of_Total'] = (
        (cat_sum['Total_Amount'] / total_amount_all * 100.0) if total_amount_all > 0 else 0.0
    ).round(1)

    # Standard order of categories
    cat_order = ['Guest Supplies', 'Cleaning Supplies', 'Paper Supplies', 'Print & Stationery']
    cat_sum['sort_order'] = cat_sum['Category'].apply(lambda c: cat_order.index(c) if c in cat_order else 99)
    cat_sum = cat_sum.sort_values('sort_order').drop(columns=['sort_order']).reset_index(drop=True)

    return cat_sum


def get_item_summary(df: pd.DataFrame, category: Optional[str] = None) -> pd.DataFrame:
    """
    Returns aggregated items summary (Total Qty, Avg Cost, Total Cost) by item.
    Optionally filtered by category.
    """
    if df.empty:
        return pd.DataFrame()

    filtered = df if category is None else df[df['Category'] == category]
    if filtered.empty:
        return pd.DataFrame()

    item_sum = filtered.groupby(['Category', 'Item_Code', 'Item_Name']).agg(
        Total_Qty=('Qty', 'sum'),
        Unit=('Unit', 'first'),
        Avg_Cost=('Cost', 'mean'),
        Total_Amount=('Total', 'sum'),
        Order_Frequency=('SR_Number', 'nunique'),
        First_Date=('Date', 'min'),
        Last_Date=('Date', 'max')
    ).reset_index()

    item_sum['Avg_Cost'] = item_sum['Avg_Cost'].round(0)
    item_sum = item_sum.sort_values(by=['Category', 'Total_Amount'], ascending=[True, False]).reset_index(drop=True)

    return item_sum
