"""exporter.py - Excel generation module for Hotel.
Creates multi-tab, beautifully styled, separated Excel workbooks using openpyxl.
"""

from typing import Dict, Any, Optional
import io
import logging
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configure module-level logger
logger = logging.getLogger(__name__)
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.INFO)


def create_separated_excel(
    reconciled_data: Dict[str, Any],
    month_name: str = "August 2026",
    output_path: Optional[str] = None
) -> bytes:
    """
    Builds a multi-tab separated Excel workbook containing:
      1. Summary & Budget Variance
      2. All Separated Transactions
      3. Guest Supplies Breakdown
      4. Cleaning Supplies Breakdown
      5. Paper Supplies Breakdown
      6. Outsourcing & Laundry Breakdown
      7. Media & Utilities Breakdown
      8. Payroll & Service Charge Breakdown
    """
    logger.info(f"Creating separated Excel for {month_name}")
    
    # Validate input data
    if not reconciled_data:
        logger.error("No reconciled data provided")
        raise ValueError("No reconciled data provided")
    
    required_keys = ['transactions', 'category_summary', 'metrics']
    for key in required_keys:
        if key not in reconciled_data:
            logger.error(f"Missing required key in reconciled_data: {key}")
            raise ValueError(f"Missing required key in reconciled_data: {key}")
    
    logger.info(f"Creating separated Excel for {month_name}")
    
    # Validate input data
    if not reconciled_data:
        logger.error("No reconciled data provided")
        raise ValueError("No reconciled data provided")
    
    required_keys = ['transactions', 'category_summary', 'metrics']
    for key in required_keys:
        if key not in reconciled_data:
            logger.error(f"Missing required key in reconciled_data: {key}")
            raise ValueError(f"Missing required key in reconciled_data: {key}")

    
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    navy_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    dark_slate_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    blue_header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    light_blue_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    
    red_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    
    white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
    subtitle_font = Font(name="Calibri", size=11, italic=True, color="64748B")
    bold_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    regular_font = Font(name="Calibri", size=10, color="0F172A")
    
    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    total_top_border = Side(style='thin', color='0F172A')
    total_bottom_border = Side(style='double', color='0F172A')
    total_border = Border(top=total_top_border, bottom=total_bottom_border)

    num_format_currency = "#,##0"
    num_format_pct = "0.0%"

    cat_df = reconciled_data['category_summary']
    trx_df = reconciled_data['transactions']
    metrics = reconciled_data['metrics']

    # ==========================================
    # SHEET 1: Summary & Budget Variance
    # ==========================================
    ws_sum = wb.create_sheet(title="Executive Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    # Title
    ws_sum["A1"] = "Hotel - Spending Breakdown & Variance"
    ws_sum["A1"].font = title_font
    ws_sum["A2"] = f"Period: {month_name} | Generated via Expense Separator App"
    ws_sum["A2"].font = subtitle_font

    # KPI Summary Cards in Rows 4-6
    kpis = [
        ("TOTAL SPENT", metrics['total_spent'], "IDR", 1),
        ("TOTAL BUDGET", metrics['total_budget'], "IDR", 3),
        ("NET VARIANCE", metrics['variance_idr'], "IDR", 5),
        ("BUDGET VARIANCE", metrics['variance_pct'] / 100.0, "PCT", 7),
    ]

    for label, val, val_type, col in kpis:
        col_letter = get_column_letter(col)
        next_col_letter = get_column_letter(col + 1)
        ws_sum.merge_cells(f"{col_letter}4:{next_col_letter}4")
        ws_sum.merge_cells(f"{col_letter}5:{next_col_letter}5")
        
        c_lbl = ws_sum[f"{col_letter}4"]
        c_lbl.value = label
        c_lbl.font = Font(name="Calibri", size=9, bold=True, color="475569")
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")
        c_lbl.fill = light_blue_fill

        c_val = ws_sum[f"{col_letter}5"]
        c_val.value = val
        c_val.font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.fill = light_blue_fill
        c_val.number_format = num_format_pct if val_type == "PCT" else num_format_currency

    # Summary Table Headers
    headers_sum = [
        "Category", "Group", "Budget (IDR)", "Actual (IDR)",
        "Variance (IDR)", "Variance (%)", "Status", "Top Cost Driver", "Transactions"
    ]
    start_row = 8
    for c_idx, h in enumerate(headers_sum, 1):
        cell = ws_sum.cell(start_row, c_idx, h)
        cell.font = white_bold
        cell.fill = dark_slate_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, (_, row) in enumerate(cat_df.iterrows(), start_row + 1):
        ws_sum.cell(r_idx, 1, row['Category']).font = regular_font
        ws_sum.cell(r_idx, 2, row['Group']).font = regular_font
        
        c_bgt = ws_sum.cell(r_idx, 3, row['Budget'])
        c_bgt.font = regular_font
        c_bgt.number_format = num_format_currency

        c_act = ws_sum.cell(r_idx, 4, row['Actual'])
        c_act.font = bold_font
        c_act.number_format = num_format_currency

        c_var = ws_sum.cell(r_idx, 5, row['Variance_IDR'])
        c_var.font = regular_font
        c_var.number_format = num_format_currency

        c_pct = ws_sum.cell(r_idx, 6, row['Variance_Pct'] / 100.0)
        c_pct.font = regular_font
        c_pct.number_format = num_format_pct

        c_stat = ws_sum.cell(r_idx, 7, row['Status'])
        c_stat.alignment = Alignment(horizontal="center")
        if row['Status'] == 'Over Budget':
            c_stat.fill = red_fill
            c_stat.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
        elif row['Status'] == 'Under Budget':
            c_stat.fill = green_fill
            c_stat.font = Font(name="Calibri", size=10, bold=True, color="166534")
        else:
            c_stat.font = regular_font

        ws_sum.cell(r_idx, 8, row['Top_Item']).font = regular_font
        ws_sum.cell(r_idx, 9, row['Transaction_Count']).font = regular_font
        ws_sum.cell(r_idx, 9).alignment = Alignment(horizontal="center")

        # Borders
        for col_i in range(1, 10):
            ws_sum.cell(r_idx, col_i).border = thin_border

    # Summary Totals Row
    if cat_df.empty:
        tot_row = start_row  # header exists but no data rows
    else:
        tot_row = start_row + len(cat_df) + 1
    first_data_row = start_row + 1
    ws_sum.cell(tot_row, 1, "TOTAL").font = bold_font
    ws_sum.cell(tot_row, 3, f"=SUM(C{first_data_row}:C{tot_row-1})").font = bold_font
    ws_sum.cell(tot_row, 3).number_format = num_format_currency
    ws_sum.cell(tot_row, 4, f"=SUM(D{first_data_row}:D{tot_row-1})").font = bold_font
    ws_sum.cell(tot_row, 4).number_format = num_format_currency
    ws_sum.cell(tot_row, 5, f"=D{tot_row}-C{tot_row}").font = bold_font
    ws_sum.cell(tot_row, 5).number_format = num_format_currency
    if tot_row > first_data_row:
        ws_sum.cell(tot_row, 6, f"=IF(C{tot_row}=0,0,E{tot_row}/C{tot_row})").font = bold_font
    ws_sum.cell(tot_row, 6).number_format = num_format_pct
    if tot_row > first_data_row:
        ws_sum.cell(tot_row, 9, f"=SUM(I{first_data_row}:I{tot_row-1})").font = bold_font
    ws_sum.cell(tot_row, 9).alignment = Alignment(horizontal="center")

    for col_i in range(1, 10):
        ws_sum.cell(tot_row, col_i).border = total_border

    # ==========================================
    # SHEET 2: All Reconciled Transactions
    # ==========================================
    ws_trx = wb.create_sheet(title="All Transactions")
    ws_trx.views.sheetView[0].showGridLines = True

    headers_trx = [
        "Date", "Category", "Account Code", "Item / Service Name",
        "Vendor / Partner", "Qty", "Unit", "Unit Price (IDR)", "Total Amount (IDR)",
        "Voucher / Ref", "JRNL", "Original Description"
    ]
    for c_idx, h in enumerate(headers_trx, 1):
        cell = ws_trx.cell(1, c_idx, h)
        cell.font = white_bold
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, (_, row) in enumerate(trx_df.iterrows(), 2):
        fill = zebra_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        ws_trx.cell(r_idx, 1, row['Date']).font = regular_font
        ws_trx.cell(r_idx, 2, row['Category']).font = regular_font
        ws_trx.cell(r_idx, 3, row.get('Account_Code', '')).font = regular_font
        ws_trx.cell(r_idx, 4, row['Item_Name']).font = bold_font
        ws_trx.cell(r_idx, 5, row['Partner_Vendor']).font = regular_font
        
        c_q = ws_trx.cell(r_idx, 6, row['Qty'])
        c_q.font = regular_font
        c_q.alignment = Alignment(horizontal="right")
        
        ws_trx.cell(r_idx, 7, row['Unit']).font = regular_font
        
        c_up = ws_trx.cell(r_idx, 8, row['Unit_Price'])
        c_up.font = regular_font
        c_up.number_format = num_format_currency

        c_amt = ws_trx.cell(r_idx, 9, row['Amount'])
        c_amt.font = bold_font
        c_amt.number_format = num_format_currency

        ws_trx.cell(r_idx, 10, row['Voucher_Ref']).font = regular_font
        ws_trx.cell(r_idx, 11, row['JRNL']).font = regular_font
        ws_trx.cell(r_idx, 12, row['Raw_Description']).font = regular_font

        for col_i in range(1, 13):
            c_cell = ws_trx.cell(r_idx, col_i)
            if fill.fill_type:
                c_cell.fill = fill
            c_cell.border = thin_border

    # ==========================================
    # SHEET 3: Item Totals & Quantities (Aggregated by Item)
    # ==========================================
    item_df = reconciled_data.get('item_summary', pd.DataFrame())
    if not item_df.empty:
        ws_items = wb.create_sheet(title="Item Totals & Quantities")
        ws_items.views.sheetView[0].showGridLines = True

        ws_items["A1"] = "Hotel - Item Consumption & Repeated Orders Summary"
        ws_items["A1"].font = title_font
        ws_items["A2"] = f"Month: {month_name} | Aggregated Total Quantities and Spending per Item"
        ws_items["A2"].font = subtitle_font

        headers_items = [
            "Category", "Item / Service Name", "Total Quantity", "Unit",
            "Avg Unit Price (IDR)", "Total Spend (IDR)", "% of Category", "Orders Count", "Date Range"
        ]
        s_row = 4
        for c_idx, h in enumerate(headers_items, 1):
            cell = ws_items.cell(s_row, c_idx, h)
            cell.font = white_bold
            cell.fill = navy_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        curr_cat = None
        r_idx = s_row + 1
        for _, row in item_df.iterrows():
            fill = zebra_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
            
            ws_items.cell(r_idx, 1, row['Category']).font = bold_font if row['Category'] != curr_cat else regular_font
            curr_cat = row['Category']
            ws_items.cell(r_idx, 2, row['Item_Name']).font = bold_font
            
            c_q = ws_items.cell(r_idx, 3, row['Total_Qty'])
            c_q.font = bold_font
            c_q.alignment = Alignment(horizontal="right")
            c_q.number_format = "#,##0"

            ws_items.cell(r_idx, 4, row['Unit']).font = regular_font
            
            c_up = ws_items.cell(r_idx, 5, row['Avg_Unit_Price'])
            c_up.font = regular_font
            c_up.number_format = num_format_currency

            c_amt = ws_items.cell(r_idx, 6, row['Total_Amount'])
            c_amt.font = bold_font
            c_amt.number_format = num_format_currency

            c_pct = ws_items.cell(r_idx, 7, row['Pct_Of_Category'] / 100.0)
            c_pct.font = regular_font
            c_pct.number_format = num_format_pct

            c_cnt = ws_items.cell(r_idx, 8, row['Order_Count'])
            c_cnt.font = regular_font
            c_cnt.alignment = Alignment(horizontal="center")

            date_range_str = f"{row['First_Date']} ~ {row['Last_Date']}" if row['First_Date'] != row['Last_Date'] else str(row['First_Date'])
            ws_items.cell(r_idx, 9, date_range_str).font = regular_font

            for col_i in range(1, 10):
                c_cell = ws_items.cell(r_idx, col_i)
                if fill.fill_type:
                    c_cell.fill = fill
                c_cell.border = thin_border

            r_idx += 1

        # Grand Total row
        ws_items.cell(r_idx, 1, "GRAND TOTAL").font = bold_font
        ws_items.cell(r_idx, 6, f"=SUM(F{s_row+1}:F{r_idx-1})").font = bold_font
        ws_items.cell(r_idx, 6).number_format = num_format_currency
        for col_i in range(1, 10):
            ws_items.cell(r_idx, col_i).border = total_border

    # ==========================================
    # HELPER: Function to create category sheets with Item Summary + Transactions
    # ==========================================
    def add_category_sheet(sheet_title: str, category_filters: list):
        filtered_df = trx_df[trx_df['Category'].isin(category_filters)].copy()
        if filtered_df.empty:
            return

        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        # Sheet header
        ws["A1"] = f"{sheet_title} - Spending Breakdown & Item Totals"
        ws["A1"].font = title_font
        ws["A2"] = f"Month: {month_name} | Total Spent: Rp {filtered_df['Amount'].sum():,.0f} across {len(filtered_df)} transactions"
        ws["A2"].font = subtitle_font

        # 1. Item Totals Sub-Table for this category
        cat_items = item_df[item_df['Category'].isin(category_filters)].copy() if not item_df.empty else pd.DataFrame()
        curr_row = 4
        
        if not cat_items.empty:
            ws.cell(curr_row, 1, "ITEM TOTALS SUMMARY (Total Quantities & Costs)").font = bold_font
            curr_row += 1
            
            headers_sub = ["Item Name", "Category", "Total Quantity", "Unit", "Avg Unit Price (IDR)", "Total Spend (IDR)", "% of Category", "Orders Count"]
            for c_idx, h in enumerate(headers_sub, 1):
                cell = ws.cell(curr_row, c_idx, h)
                cell.font = white_bold
                cell.fill = blue_header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            curr_row += 1
            start_sub_row = curr_row
            for _, i_row in cat_items.iterrows():
                ws.cell(curr_row, 1, i_row['Item_Name']).font = bold_font
                ws.cell(curr_row, 2, i_row['Category']).font = regular_font
                
                c_q = ws.cell(curr_row, 3, i_row['Total_Qty'])
                c_q.font = bold_font
                c_q.alignment = Alignment(horizontal="right")
                c_q.number_format = "#,##0"

                ws.cell(curr_row, 4, i_row['Unit']).font = regular_font
                
                c_up = ws.cell(curr_row, 5, i_row['Avg_Unit_Price'])
                c_up.font = regular_font
                c_up.number_format = num_format_currency

                c_amt = ws.cell(curr_row, 6, i_row['Total_Amount'])
                c_amt.font = bold_font
                c_amt.number_format = num_format_currency

                c_pct = ws.cell(curr_row, 7, i_row['Pct_Of_Category'] / 100.0)
                c_pct.font = regular_font
                c_pct.number_format = num_format_pct

                c_cnt = ws.cell(curr_row, 8, i_row['Order_Count'])
                c_cnt.font = regular_font
                c_cnt.alignment = Alignment(horizontal="center")

                for col_i in range(1, 9):
                    ws.cell(curr_row, col_i).border = thin_border
                curr_row += 1

            # Subtotal of Item Summary
            ws.cell(curr_row, 1, "ITEMS SUB-TOTAL").font = bold_font
            ws.cell(curr_row, 6, f"=SUM(F{start_sub_row}:F{curr_row-1})").font = bold_font
            ws.cell(curr_row, 6).number_format = num_format_currency
            for col_i in range(1, 9):
                ws.cell(curr_row, col_i).border = total_border

            curr_row += 2  # spacing before transaction log

        # 2. Individual Transaction Log
        ws.cell(curr_row, 1, "INDIVIDUAL TRANSACTIONS (Order / Issuing Log)").font = bold_font
        curr_row += 1

        headers = [
            "Date", "Category", "Item Name", "Vendor / Source",
            "Qty", "Unit", "Unit Price", "Total Amount (IDR)", "Voucher / Ref"
        ]
        s_row = curr_row
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(s_row, c_idx, h)
            cell.font = white_bold
            cell.fill = dark_slate_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_idx, (_, row) in enumerate(filtered_df.iterrows(), s_row + 1):
            ws.cell(r_idx, 1, row['Date']).font = regular_font
            ws.cell(r_idx, 2, row['Category']).font = regular_font
            ws.cell(r_idx, 3, row['Item_Name']).font = bold_font
            ws.cell(r_idx, 4, row['Partner_Vendor']).font = regular_font
            ws.cell(r_idx, 5, row['Qty']).font = regular_font
            ws.cell(r_idx, 6, row['Unit']).font = regular_font
            
            c_p = ws.cell(r_idx, 7, row['Unit_Price'])
            c_p.font = regular_font
            c_p.number_format = num_format_currency

            c_a = ws.cell(r_idx, 8, row['Amount'])
            c_a.font = bold_font
            c_a.number_format = num_format_currency

            ws.cell(r_idx, 9, row['Voucher_Ref']).font = regular_font

            for col_i in range(1, 10):
                ws.cell(r_idx, col_i).border = thin_border

        # Category Total
        end_row = s_row + len(filtered_df) + 1
        ws.cell(end_row, 1, "TRANSACTION TOTAL").font = bold_font
        ws.cell(end_row, 8, f"=SUM(H{s_row+1}:H{end_row-1})").font = bold_font
        ws.cell(end_row, 8).number_format = num_format_currency
        for col_i in range(1, 10):
            ws.cell(end_row, col_i).border = total_border

    # Generate dedicated separated category tabs
    add_category_sheet("Guest Supplies", ["Guest Supplies"])
    add_category_sheet("Outsourcing & Laundry", ["Outsourcing Utilities", "Laundry & Dry Cleaning"])
    add_category_sheet("Cleaning Supplies", ["Cleaning Supplies"])
    add_category_sheet("Paper Supplies", ["Paper Suplies"])
    add_category_sheet("Payroll & SC", [
        "Salaries & Wages", "Salaries & Wages (kary.lepas)", "Service Charge (Account Perampungan PPh)",
        "Function Allowance", "Employess Transportation", "Human Resources", "PTEB"
    ])
    add_category_sheet("Media & Utilities", [
        "Music Cassette Entertainment & Tv Kabel", "Telephone &  Internet"
    ])

    # Auto-adjust column widths across all sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None and not str(val).startswith("="):
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    if output_path:
        wb.save(output_path)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def create_sr_separated_excel(
    sr_df: pd.DataFrame,
    report_title: str = "Stock Request Separated Report",
    output_path: Optional[str] = None
) -> bytes:
    """
    Builds a multi-tab separated Excel workbook for Stock Requests:
      1. SR Executive Summary
      2. Guest Supplies (Item Summary + Issuing Log)
      3. Cleaning Supplies (Item Summary + Issuing Log)
      4. Paper Supplies (Item Summary + Issuing Log)
      5. Print & Stationery (Item Summary + Issuing Log)
      6. All Stock Requests Log
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Styles
    navy_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    dark_slate_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    blue_header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    light_blue_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
    subtitle_font = Font(name="Calibri", size=11, italic=True, color="64748B")
    bold_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    regular_font = Font(name="Calibri", size=10, color="0F172A")

    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    total_top_border = Side(style='thin', color='0F172A')
    total_bottom_border = Side(style='double', color='0F172A')
    total_border = Border(top=total_top_border, bottom=total_bottom_border)

    num_format_currency = "#,##0"
    num_format_pct = "0.0%"

    total_amount = sr_df['Total'].sum() if not sr_df.empty else 0.0
    total_qty = sr_df['Qty'].sum() if not sr_df.empty else 0.0
    total_items = len(sr_df) if not sr_df.empty else 0
    unique_srs = sr_df['SR_Number'].nunique() if not sr_df.empty else 0

    # ==========================================
    # SHEET 1: SR Executive Summary
    # ==========================================
    ws_sum = wb.create_sheet(title="SR Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum["A1"] = "Hotel - Stock Request (SR) Separator"
    ws_sum["A1"].font = title_font
    ws_sum["A2"] = f"{report_title} | Gudang Central Issuing to Housekeeping"
    ws_sum["A2"].font = subtitle_font

    # KPI Summary Cards
    kpis = [
        ("TOTAL SPEND", total_amount, "IDR", 1),
        ("TOTAL QUANTITY", total_qty, "QTY", 3),
        ("LINE ITEMS", total_items, "COUNT", 5),
        ("SR ORDERS", unique_srs, "COUNT", 7),
    ]

    for label, val, val_type, col in kpis:
        col_letter = get_column_letter(col)
        next_col_letter = get_column_letter(col + 1)
        ws_sum.merge_cells(f"{col_letter}4:{next_col_letter}4")
        ws_sum.merge_cells(f"{col_letter}5:{next_col_letter}5")

        c_lbl = ws_sum[f"{col_letter}4"]
        c_lbl.value = label
        c_lbl.font = Font(name="Calibri", size=9, bold=True, color="475569")
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")
        c_lbl.fill = light_blue_fill

        c_val = ws_sum[f"{col_letter}5"]
        c_val.value = val
        c_val.font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.fill = light_blue_fill
        c_val.number_format = num_format_currency if val_type in ("IDR", "QTY") else "#,##0"

    # Category Summary Table
    ws_sum.cell(7, 1, "SEPARATED SPENDING BY CATEGORY").font = bold_font

    headers_cat = ["Category", "Unique Items", "Total Quantity", "Total Amount (IDR)", "% of Total Spend", "Issuing Count"]
    for c_idx, h in enumerate(headers_cat, 1):
        cell = ws_sum.cell(8, c_idx, h)
        cell.font = white_bold
        cell.fill = dark_slate_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cat_order = ['Guest Supplies', 'Cleaning Supplies', 'Paper Supplies', 'Print & Stationery']
    curr_r = 9
    for cat in cat_order:
        cat_data = sr_df[sr_df['Category'] == cat] if not sr_df.empty else pd.DataFrame()
        c_qty = cat_data['Qty'].sum() if not cat_data.empty else 0.0
        c_amt = cat_data['Total'].sum() if not cat_data.empty else 0.0
        c_items = cat_data['Item_Name'].nunique() if not cat_data.empty else 0
        c_count = len(cat_data) if not cat_data.empty else 0
        c_pct = (c_amt / total_amount) if total_amount > 0 else 0.0

        ws_sum.cell(curr_r, 1, cat).font = bold_font
        ws_sum.cell(curr_r, 2, c_items).font = regular_font
        ws_sum.cell(curr_r, 2).alignment = Alignment(horizontal="center")

        ws_sum.cell(curr_r, 3, c_qty).font = regular_font
        ws_sum.cell(curr_r, 3).number_format = "#,##0"

        ws_sum.cell(curr_r, 4, c_amt).font = bold_font
        ws_sum.cell(curr_r, 4).number_format = num_format_currency

        ws_sum.cell(curr_r, 5, c_pct).font = regular_font
        ws_sum.cell(curr_r, 5).number_format = num_format_pct

        ws_sum.cell(curr_r, 6, c_count).font = regular_font
        ws_sum.cell(curr_r, 6).alignment = Alignment(horizontal="center")

        for col_i in range(1, 7):
            ws_sum.cell(curr_r, col_i).border = thin_border
        curr_r += 1

    # Total row
    ws_sum.cell(curr_r, 1, "TOTAL").font = bold_font
    ws_sum.cell(curr_r, 3, f"=SUM(C9:C{curr_r-1})").font = bold_font
    ws_sum.cell(curr_r, 3).number_format = "#,##0"
    ws_sum.cell(curr_r, 4, f"=SUM(D9:D{curr_r-1})").font = bold_font
    ws_sum.cell(curr_r, 4).number_format = num_format_currency
    ws_sum.cell(curr_r, 5, 1.0).font = bold_font
    ws_sum.cell(curr_r, 5).number_format = num_format_pct
    ws_sum.cell(curr_r, 6, f"=SUM(F9:F{curr_r-1})").font = bold_font
    ws_sum.cell(curr_r, 6).alignment = Alignment(horizontal="center")

    for col_i in range(1, 7):
        ws_sum.cell(curr_r, col_i).border = total_border

    # Top 10 Cost Drivers
    curr_r += 2
    ws_sum.cell(curr_r, 1, "TOP 10 COST DRIVERS ACROSS ALL STOCK REQUESTS").font = bold_font
    curr_r += 1

    headers_top = ["Rank", "Item Name", "Category", "Total Qty", "Unit", "Total Cost (IDR)", "% of Spend"]
    for c_idx, h in enumerate(headers_top, 1):
        cell = ws_sum.cell(curr_r, c_idx, h)
        cell.font = white_bold
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    curr_r += 1

    if not sr_df.empty:
        top_items = sr_df.groupby(['Item_Name', 'Category', 'Unit'], as_index=False).agg(
            Total_Qty=('Qty', 'sum'),
            Total_Amount=('Total', 'sum')
        ).sort_values(by='Total_Amount', ascending=False).head(10).reset_index(drop=True)

        for rank, (_, row) in enumerate(top_items.iterrows(), 1):
            ws_sum.cell(curr_r, 1, rank).font = regular_font
            ws_sum.cell(curr_r, 1).alignment = Alignment(horizontal="center")
            ws_sum.cell(curr_r, 2, row['Item_Name']).font = bold_font
            ws_sum.cell(curr_r, 3, row['Category']).font = regular_font
            ws_sum.cell(curr_r, 4, row['Total_Qty']).font = regular_font
            ws_sum.cell(curr_r, 4).number_format = "#,##0"
            ws_sum.cell(curr_r, 5, row['Unit']).font = regular_font
            ws_sum.cell(curr_r, 6, row['Total_Amount']).font = bold_font
            ws_sum.cell(curr_r, 6).number_format = num_format_currency
            pct = row['Total_Amount'] / total_amount if total_amount > 0 else 0.0
            ws_sum.cell(curr_r, 7, pct).font = regular_font
            ws_sum.cell(curr_r, 7).number_format = num_format_pct

            for col_i in range(1, 8):
                ws_sum.cell(curr_r, col_i).border = thin_border
            curr_r += 1

    # ==========================================
    # HELPER: Function to create a category tab
    # ==========================================
    def add_sr_category_tab(cat_name: str):
        cat_df = sr_df[sr_df['Category'] == cat_name].copy() if not sr_df.empty else pd.DataFrame()
        if cat_df.empty:
            return

        ws = wb.create_sheet(title=cat_name)
        ws.views.sheetView[0].showGridLines = True

        cat_total = cat_df['Total'].sum()
        cat_qty = cat_df['Qty'].sum()

        ws["A1"] = f"{cat_name} - Stock Request Breakdown"
        ws["A1"].font = title_font
        ws["A2"] = f"Total Spend: Rp {cat_total:,.0f} | Total Quantity: {cat_qty:,.0f} across {len(cat_df)} requests"
        ws["A2"].font = subtitle_font

        # 1. Aggregated Item Summary
        ws.cell(4, 1, "ITEM SUMMARY (Total Quantities & Costs)").font = bold_font

        item_agg = cat_df.groupby(['Item_Code', 'Item_Name', 'Unit'], as_index=False).agg(
            Total_Qty=('Qty', 'sum'),
            Avg_Cost=('Cost', 'mean'),
            Total_Amount=('Total', 'sum'),
            Order_Count=('SR_Number', 'nunique')
        ).sort_values(by='Total_Amount', ascending=False).reset_index(drop=True)

        headers_item = ["Item Code", "Item Name", "Total Qty", "Unit", "Avg Unit Cost (IDR)", "Total Spend (IDR)", "% of Category", "Orders Count"]
        for c_idx, h in enumerate(headers_item, 1):
            cell = ws.cell(5, c_idx, h)
            cell.font = white_bold
            cell.fill = blue_header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        r = 6
        start_item_r = r
        for _, row in item_agg.iterrows():
            ws.cell(r, 1, row['Item_Code']).font = regular_font
            ws.cell(r, 2, row['Item_Name']).font = bold_font
            ws.cell(r, 3, row['Total_Qty']).font = bold_font
            ws.cell(r, 3).number_format = "#,##0"
            ws.cell(r, 3).alignment = Alignment(horizontal="right")
            ws.cell(r, 4, row['Unit']).font = regular_font
            ws.cell(r, 5, row['Avg_Cost']).font = regular_font
            ws.cell(r, 5).number_format = num_format_currency
            ws.cell(r, 6, row['Total_Amount']).font = bold_font
            ws.cell(r, 6).number_format = num_format_currency
            pct = (row['Total_Amount'] / cat_total) if cat_total > 0 else 0.0
            ws.cell(r, 7, pct).font = regular_font
            ws.cell(r, 7).number_format = num_format_pct
            ws.cell(r, 8, row['Order_Count']).font = regular_font
            ws.cell(r, 8).alignment = Alignment(horizontal="center")

            for col_i in range(1, 9):
                ws.cell(r, col_i).border = thin_border
            r += 1

        # Item Summary Subtotal
        ws.cell(r, 2, "SUBTOTAL").font = bold_font
        ws.cell(r, 3, f"=SUM(C{start_item_r}:C{r-1})").font = bold_font
        ws.cell(r, 3).number_format = "#,##0"
        ws.cell(r, 6, f"=SUM(F{start_item_r}:F{r-1})").font = bold_font
        ws.cell(r, 6).number_format = num_format_currency
        for col_i in range(1, 9):
            ws.cell(r, col_i).border = total_border

        # 2. Detailed Issuing History (Chronological)
        r += 2
        ws.cell(r, 1, "CHRONOLOGICAL STOCK REQUEST / ISSUING LOG").font = bold_font
        r += 1

        headers_log = ["Date", "SR Reference", "Item Code", "Item Name", "Qty", "Unit", "Unit Cost (IDR)", "Total (IDR)", "Requested By"]
        for c_idx, h in enumerate(headers_log, 1):
            cell = ws.cell(r, c_idx, h)
            cell.font = white_bold
            cell.fill = dark_slate_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        r += 1

        start_log_r = r
        cat_sorted = cat_df.sort_values(by=['Date', 'SR_Number', 'Item_Name'])
        for _, row in cat_sorted.iterrows():
            ws.cell(r, 1, row['Date']).font = regular_font
            ws.cell(r, 2, row['SR_Number']).font = regular_font
            ws.cell(r, 3, row['Item_Code']).font = regular_font
            ws.cell(r, 4, row['Item_Name']).font = bold_font
            ws.cell(r, 5, row['Qty']).font = regular_font
            ws.cell(r, 5).number_format = "#,##0"
            ws.cell(r, 6, row['Unit']).font = regular_font
            ws.cell(r, 7, row['Cost']).font = regular_font
            ws.cell(r, 7).number_format = num_format_currency
            ws.cell(r, 8, row['Total']).font = bold_font
            ws.cell(r, 8).number_format = num_format_currency
            ws.cell(r, 9, row['Requested_By']).font = regular_font

            for col_i in range(1, 10):
                ws.cell(r, col_i).border = thin_border
            r += 1

        # Log Subtotal
        ws.cell(r, 4, "TOTAL ISSUED").font = bold_font
        ws.cell(r, 5, f"=SUM(E{start_log_r}:E{r-1})").font = bold_font
        ws.cell(r, 5).number_format = "#,##0"
        ws.cell(r, 8, f"=SUM(H{start_log_r}:H{r-1})").font = bold_font
        ws.cell(r, 8).number_format = num_format_currency
        for col_i in range(1, 10):
            ws.cell(r, col_i).border = total_border

    # Create category tabs
    add_sr_category_tab("Guest Supplies")
    add_sr_category_tab("Cleaning Supplies")
    add_sr_category_tab("Paper Supplies")
    add_sr_category_tab("Print & Stationery")

    # ==========================================
    # SHEET 6: All Stock Requests Log
    # ==========================================
    ws_all = wb.create_sheet(title="All SR Items")
    ws_all.views.sheetView[0].showGridLines = True

    headers_all = [
        "Date", "SR Reference", "Category", "Item Code", "Item Name",
        "Qty", "Unit", "Unit Cost (IDR)", "Total (IDR)", "Requested By", "Source File"
    ]
    for c_idx, h in enumerate(headers_all, 1):
        cell = ws_all.cell(1, c_idx, h)
        cell.font = white_bold
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, (_, row) in enumerate(sr_df.iterrows(), 2):
        ws_all.cell(r_idx, 1, row['Date']).font = regular_font
        ws_all.cell(r_idx, 2, row['SR_Number']).font = regular_font
        ws_all.cell(r_idx, 3, row['Category']).font = bold_font
        ws_all.cell(r_idx, 4, row['Item_Code']).font = regular_font
        ws_all.cell(r_idx, 5, row['Item_Name']).font = bold_font
        ws_all.cell(r_idx, 6, row['Qty']).font = regular_font
        ws_all.cell(r_idx, 6).number_format = "#,##0"
        ws_all.cell(r_idx, 7, row['Unit']).font = regular_font
        ws_all.cell(r_idx, 8, row['Cost']).font = regular_font
        ws_all.cell(r_idx, 8).number_format = num_format_currency
        ws_all.cell(r_idx, 9, row['Total']).font = bold_font
        ws_all.cell(r_idx, 9).number_format = num_format_currency
        ws_all.cell(r_idx, 10, row['Requested_By']).font = regular_font
        ws_all.cell(r_idx, 11, row.get('Source_File', '')).font = regular_font

        for col_i in range(1, 12):
            ws_all.cell(r_idx, col_i).border = thin_border

    # Auto-adjust column widths across all sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None and not str(val).startswith("="):
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    if output_path:
        wb.save(output_path)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

