"""exporter.py - Excel generation module for Hotel.
Creates multi-tab, beautifully styled, separated Excel workbooks using openpyxl.
"""

from typing import Dict, Any, Optional
import io
import logging
from datetime import datetime, timedelta
import pandas as pd
import openpyxl

from ui.security import sanitize_excel_value
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Reusable number formats for all exported sheets
NUM_FMT_CURRENCY = "#,##0"
NUM_FMT_PCT = "0.0%"

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
    month_name = sanitize_excel_value(month_name)
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
        c_val.number_format = NUM_FMT_PCT if val_type == "PCT" else NUM_FMT_CURRENCY

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
        ws_sum.cell(r_idx, 1, sanitize_excel_value(row['Category'])).font = regular_font
        ws_sum.cell(r_idx, 2, sanitize_excel_value(row['Group'])).font = regular_font
        
        c_bgt = ws_sum.cell(r_idx, 3, row['Budget'])
        c_bgt.font = regular_font
        c_bgt.number_format = NUM_FMT_CURRENCY

        c_act = ws_sum.cell(r_idx, 4, row['Actual'])
        c_act.font = bold_font
        c_act.number_format = NUM_FMT_CURRENCY

        c_var = ws_sum.cell(r_idx, 5, row['Variance_IDR'])
        c_var.font = regular_font
        c_var.number_format = NUM_FMT_CURRENCY

        c_pct = ws_sum.cell(r_idx, 6, row['Variance_Pct'] / 100.0)
        c_pct.font = regular_font
        c_pct.number_format = NUM_FMT_PCT

        c_stat = ws_sum.cell(r_idx, 7, sanitize_excel_value(row['Status']))
        c_stat.alignment = Alignment(horizontal="center")
        if row['Status'] == 'Over Budget':
            c_stat.fill = red_fill
            c_stat.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
        elif row['Status'] == 'Under Budget':
            c_stat.fill = green_fill
            c_stat.font = Font(name="Calibri", size=10, bold=True, color="166534")
        else:
            c_stat.font = regular_font

        ws_sum.cell(r_idx, 8, sanitize_excel_value(row['Top_Item'])).font = regular_font
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
    ws_sum.cell(tot_row, 3).number_format = NUM_FMT_CURRENCY
    ws_sum.cell(tot_row, 4, f"=SUM(D{first_data_row}:D{tot_row-1})").font = bold_font
    ws_sum.cell(tot_row, 4).number_format = NUM_FMT_CURRENCY
    ws_sum.cell(tot_row, 5, f"=D{tot_row}-C{tot_row}").font = bold_font
    ws_sum.cell(tot_row, 5).number_format = NUM_FMT_CURRENCY
    if tot_row > first_data_row:
        ws_sum.cell(tot_row, 6, f"=IF(C{tot_row}=0,0,E{tot_row}/C{tot_row})").font = bold_font
    ws_sum.cell(tot_row, 6).number_format = NUM_FMT_PCT
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
        ws_trx.cell(r_idx, 2, sanitize_excel_value(row['Category'])).font = regular_font
        ws_trx.cell(r_idx, 3, sanitize_excel_value(row.get('Account_Code', ''))).font = regular_font
        ws_trx.cell(r_idx, 4, sanitize_excel_value(row['Item_Name'])).font = bold_font
        ws_trx.cell(r_idx, 5, sanitize_excel_value(row['Partner_Vendor'])).font = regular_font
        
        c_q = ws_trx.cell(r_idx, 6, row['Qty'])
        c_q.font = regular_font
        c_q.alignment = Alignment(horizontal="right")
        
        ws_trx.cell(r_idx, 7, sanitize_excel_value(row['Unit'])).font = regular_font
        
        c_up = ws_trx.cell(r_idx, 8, row['Unit_Price'])
        c_up.font = regular_font
        c_up.number_format = NUM_FMT_CURRENCY

        c_amt = ws_trx.cell(r_idx, 9, row['Amount'])
        c_amt.font = bold_font
        c_amt.number_format = NUM_FMT_CURRENCY

        ws_trx.cell(r_idx, 10, sanitize_excel_value(row['Voucher_Ref'])).font = regular_font
        ws_trx.cell(r_idx, 11, sanitize_excel_value(row['JRNL'])).font = regular_font
        ws_trx.cell(r_idx, 12, sanitize_excel_value(row['Raw_Description'])).font = regular_font

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
            
            ws_items.cell(r_idx, 1, sanitize_excel_value(row['Category'])).font = bold_font if row['Category'] != curr_cat else regular_font
            curr_cat = row['Category']
            ws_items.cell(r_idx, 2, sanitize_excel_value(row['Item_Name'])).font = bold_font
            
            c_q = ws_items.cell(r_idx, 3, row['Total_Qty'])
            c_q.font = bold_font
            c_q.alignment = Alignment(horizontal="right")
            c_q.number_format = "#,##0"

            ws_items.cell(r_idx, 4, sanitize_excel_value(row['Unit'])).font = regular_font
            
            c_up = ws_items.cell(r_idx, 5, row['Avg_Unit_Price'])
            c_up.font = regular_font
            c_up.number_format = NUM_FMT_CURRENCY

            c_amt = ws_items.cell(r_idx, 6, row['Total_Amount'])
            c_amt.font = bold_font
            c_amt.number_format = NUM_FMT_CURRENCY

            c_pct = ws_items.cell(r_idx, 7, row['Pct_Of_Category'] / 100.0)
            c_pct.font = regular_font
            c_pct.number_format = NUM_FMT_PCT

            c_cnt = ws_items.cell(r_idx, 8, row['Order_Count'])
            c_cnt.font = regular_font
            c_cnt.alignment = Alignment(horizontal="center")

            date_range_str = f"{row['First_Date']} ~ {row['Last_Date']}" if row['First_Date'] != row['Last_Date'] else str(row['First_Date'])
            ws_items.cell(r_idx, 9, sanitize_excel_value(date_range_str)).font = regular_font

            for col_i in range(1, 10):
                c_cell = ws_items.cell(r_idx, col_i)
                if fill.fill_type:
                    c_cell.fill = fill
                c_cell.border = thin_border

            r_idx += 1

        # Grand Total row
        ws_items.cell(r_idx, 1, "GRAND TOTAL").font = bold_font
        ws_items.cell(r_idx, 6, f"=SUM(F{s_row+1}:F{r_idx-1})").font = bold_font
        ws_items.cell(r_idx, 6).number_format = NUM_FMT_CURRENCY
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
                ws.cell(curr_row, 1, sanitize_excel_value(i_row['Item_Name'])).font = bold_font
                ws.cell(curr_row, 2, sanitize_excel_value(i_row['Category'])).font = regular_font
                
                c_q = ws.cell(curr_row, 3, i_row['Total_Qty'])
                c_q.font = bold_font
                c_q.alignment = Alignment(horizontal="right")
                c_q.number_format = "#,##0"

                ws.cell(curr_row, 4, sanitize_excel_value(i_row['Unit'])).font = regular_font
                
                c_up = ws.cell(curr_row, 5, i_row['Avg_Unit_Price'])
                c_up.font = regular_font
                c_up.number_format = NUM_FMT_CURRENCY

                c_amt = ws.cell(curr_row, 6, i_row['Total_Amount'])
                c_amt.font = bold_font
                c_amt.number_format = NUM_FMT_CURRENCY

                c_pct = ws.cell(curr_row, 7, i_row['Pct_Of_Category'] / 100.0)
                c_pct.font = regular_font
                c_pct.number_format = NUM_FMT_PCT

                c_cnt = ws.cell(curr_row, 8, i_row['Order_Count'])
                c_cnt.font = regular_font
                c_cnt.alignment = Alignment(horizontal="center")

                for col_i in range(1, 9):
                    ws.cell(curr_row, col_i).border = thin_border
                curr_row += 1

            # Subtotal of Item Summary
            ws.cell(curr_row, 1, "ITEMS SUB-TOTAL").font = bold_font
            ws.cell(curr_row, 6, f"=SUM(F{start_sub_row}:F{curr_row-1})").font = bold_font
            ws.cell(curr_row, 6).number_format = NUM_FMT_CURRENCY
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
            ws.cell(r_idx, 2, sanitize_excel_value(row['Category'])).font = regular_font
            ws.cell(r_idx, 3, sanitize_excel_value(row['Item_Name'])).font = bold_font
            ws.cell(r_idx, 4, sanitize_excel_value(row['Partner_Vendor'])).font = regular_font
            ws.cell(r_idx, 5, row['Qty']).font = regular_font
            ws.cell(r_idx, 6, sanitize_excel_value(row['Unit'])).font = regular_font
            
            c_p = ws.cell(r_idx, 7, row['Unit_Price'])
            c_p.font = regular_font
            c_p.number_format = NUM_FMT_CURRENCY

            c_a = ws.cell(r_idx, 8, row['Amount'])
            c_a.font = bold_font
            c_a.number_format = NUM_FMT_CURRENCY

            ws.cell(r_idx, 9, sanitize_excel_value(row['Voucher_Ref'])).font = regular_font

            for col_i in range(1, 10):
                ws.cell(r_idx, col_i).border = thin_border

        # Category Total
        end_row = s_row + len(filtered_df) + 1
        ws.cell(end_row, 1, "TRANSACTION TOTAL").font = bold_font
        ws.cell(end_row, 8, f"=SUM(H{s_row+1}:H{end_row-1})").font = bold_font
        ws.cell(end_row, 8).number_format = NUM_FMT_CURRENCY
        for col_i in range(1, 10):
            ws.cell(end_row, col_i).border = total_border

    # Generate dedicated separated category tabs
    add_category_sheet("Guest Supplies", ["Guest Supplies"])
    add_category_sheet("Outsourcing & Laundry", ["Outsourcing Utilities", "Laundry & Dry Cleaning"])
    add_category_sheet("Cleaning Supplies", ["Cleaning Supplies"])
    add_category_sheet("Paper Supplies", ["Paper Supplies"])
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
    output_path: Optional[str] = None,
    occupancy_pct: Optional[float] = None,
    total_rooms_sold: Optional[int] = None,
    room_cost_daily_df: Optional[pd.DataFrame] = None,
    room_cost_monthly_df: Optional[pd.DataFrame] = None,
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


    total_amount = sr_df['Total'].sum() if not sr_df.empty else 0.0
    total_qty = sr_df['Qty'].sum() if not sr_df.empty else 0.0
    total_items = len(sr_df) if not sr_df.empty else 0
    unique_srs = sr_df['SR_Number'].nunique() if not sr_df.empty else 0
    rooms_sold_value = total_rooms_sold if total_rooms_sold is not None else 0
    rooms_sold_valid = rooms_sold_value > 0
    total_spend_per_room_sold = total_amount / rooms_sold_value if rooms_sold_valid else "N/A"

    # ==========================================
    # SHEET 1: SR Executive Summary
    # ==========================================
    ws_sum = wb.create_sheet(title="SR Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum["A1"] = "Hotel - Stock Request (SR) Separator"
    ws_sum["A1"].font = title_font
    ws_sum["A2"] = f"{sanitize_excel_value(report_title)} | Gudang Central Issuing to Housekeeping"
    ws_sum["A2"].font = subtitle_font

    # KPI Summary Cards
    kpis = [
        ("TOTAL SPEND", total_amount, "IDR", 1),
        ("TOTAL QUANTITY", total_qty, "QTY", 3),
        ("LINE ITEMS", total_items, "COUNT", 5),
        ("SR ORDERS", unique_srs, "COUNT", 7),
        ("OCCUPANCY", occupancy_pct / 100.0 if occupancy_pct is not None else "N/A", "PCT", 9),
        ("ROOMS SOLD", total_rooms_sold if total_rooms_sold is not None else "N/A", "COUNT", 11),
        ("SR SPEND / ROOM SOLD", total_spend_per_room_sold, "IDR", 13),
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
        if val_type in ("IDR", "QTY"):
            c_val.number_format = NUM_FMT_CURRENCY
        elif val_type == "PCT":
            c_val.number_format = NUM_FMT_PCT
        else:
            c_val.number_format = "#,##0"

    # Category Summary Table
    ws_sum.cell(7, 1, "SEPARATED SPENDING BY CATEGORY").font = bold_font

    headers_cat = ["Category", "Unique Items", "Total Quantity", "Total Amount (IDR)", "% of Total Spend", "Issuing Count", "Spend / Room Sold (IDR)"]
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
        c_spend_per_room_sold = c_amt / rooms_sold_value if rooms_sold_valid else "N/A"

        ws_sum.cell(curr_r, 1, cat).font = bold_font
        ws_sum.cell(curr_r, 2, c_items).font = regular_font
        ws_sum.cell(curr_r, 2).alignment = Alignment(horizontal="center")

        ws_sum.cell(curr_r, 3, c_qty).font = regular_font
        ws_sum.cell(curr_r, 3).number_format = "#,##0"

        ws_sum.cell(curr_r, 4, c_amt).font = bold_font
        ws_sum.cell(curr_r, 4).number_format = NUM_FMT_CURRENCY

        ws_sum.cell(curr_r, 5, c_pct).font = regular_font
        ws_sum.cell(curr_r, 5).number_format = NUM_FMT_PCT

        ws_sum.cell(curr_r, 6, c_count).font = regular_font
        ws_sum.cell(curr_r, 6).alignment = Alignment(horizontal="center")
        ws_sum.cell(curr_r, 7, c_spend_per_room_sold).font = regular_font
        ws_sum.cell(curr_r, 7).number_format = NUM_FMT_CURRENCY

        for col_i in range(1, 8):
            ws_sum.cell(curr_r, col_i).border = thin_border
        curr_r += 1

    # Total row
    ws_sum.cell(curr_r, 1, "TOTAL").font = bold_font
    ws_sum.cell(curr_r, 3, f"=SUM(C9:C{curr_r-1})").font = bold_font
    ws_sum.cell(curr_r, 3).number_format = "#,##0"
    ws_sum.cell(curr_r, 4, f"=SUM(D9:D{curr_r-1})").font = bold_font
    ws_sum.cell(curr_r, 4).number_format = NUM_FMT_CURRENCY
    ws_sum.cell(curr_r, 5, 1.0).font = bold_font
    ws_sum.cell(curr_r, 5).number_format = NUM_FMT_PCT
    ws_sum.cell(curr_r, 6, f"=SUM(F9:F{curr_r-1})").font = bold_font
    ws_sum.cell(curr_r, 6).alignment = Alignment(horizontal="center")
    ws_sum.cell(curr_r, 7, total_spend_per_room_sold).font = bold_font
    ws_sum.cell(curr_r, 7).number_format = NUM_FMT_CURRENCY

    for col_i in range(1, 8):
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
            ws_sum.cell(curr_r, 2, sanitize_excel_value(row['Item_Name'])).font = bold_font
            ws_sum.cell(curr_r, 3, sanitize_excel_value(row['Category'])).font = regular_font
            ws_sum.cell(curr_r, 4, row['Total_Qty']).font = regular_font
            ws_sum.cell(curr_r, 4).number_format = "#,##0"
            ws_sum.cell(curr_r, 5, sanitize_excel_value(row['Unit'])).font = regular_font
            ws_sum.cell(curr_r, 6, row['Total_Amount']).font = bold_font
            ws_sum.cell(curr_r, 6).number_format = NUM_FMT_CURRENCY
            pct = row['Total_Amount'] / total_amount if total_amount > 0 else 0.0
            ws_sum.cell(curr_r, 7, pct).font = regular_font
            ws_sum.cell(curr_r, 7).number_format = NUM_FMT_PCT

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

        headers_item = ["Item Code", "Item Name", "Total Qty", "Unit", "Avg Unit Cost (IDR)", "Total Spend (IDR)", "% of Category", "Orders Count", "Qty / Room Sold"]
        for c_idx, h in enumerate(headers_item, 1):
            cell = ws.cell(5, c_idx, h)
            cell.font = white_bold
            cell.fill = blue_header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        r = 6
        start_item_r = r
        for _, row in item_agg.iterrows():
            ws.cell(r, 1, sanitize_excel_value(row['Item_Code'])).font = regular_font
            ws.cell(r, 2, sanitize_excel_value(row['Item_Name'])).font = bold_font
            ws.cell(r, 3, row['Total_Qty']).font = bold_font
            ws.cell(r, 3).number_format = "#,##0"
            ws.cell(r, 3).alignment = Alignment(horizontal="right")
            ws.cell(r, 4, sanitize_excel_value(row['Unit'])).font = regular_font
            ws.cell(r, 5, row['Avg_Cost']).font = regular_font
            ws.cell(r, 5).number_format = NUM_FMT_CURRENCY
            ws.cell(r, 6, row['Total_Amount']).font = bold_font
            ws.cell(r, 6).number_format = NUM_FMT_CURRENCY
            pct = (row['Total_Amount'] / cat_total) if cat_total > 0 else 0.0
            ws.cell(r, 7, pct).font = regular_font
            ws.cell(r, 7).number_format = NUM_FMT_PCT
            ws.cell(r, 8, row['Order_Count']).font = regular_font
            ws.cell(r, 8).alignment = Alignment(horizontal="center")
            qty_per_room_sold = row['Total_Qty'] / rooms_sold_value if rooms_sold_valid else "N/A"
            ws.cell(r, 9, qty_per_room_sold).font = regular_font
            ws.cell(r, 9).number_format = "#,##0.0000"

            for col_i in range(1, 10):
                ws.cell(r, col_i).border = thin_border
            r += 1

        # Item Summary Subtotal
        ws.cell(r, 2, "SUBTOTAL").font = bold_font
        ws.cell(r, 3, f"=SUM(C{start_item_r}:C{r-1})").font = bold_font
        ws.cell(r, 3).number_format = "#,##0"
        ws.cell(r, 6, f"=SUM(F{start_item_r}:F{r-1})").font = bold_font
        ws.cell(r, 6).number_format = NUM_FMT_CURRENCY
        for col_i in range(1, 10):
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
            ws.cell(r, 2, sanitize_excel_value(row['SR_Number'])).font = regular_font
            ws.cell(r, 3, sanitize_excel_value(row['Item_Code'])).font = regular_font
            ws.cell(r, 4, sanitize_excel_value(row['Item_Name'])).font = bold_font
            ws.cell(r, 5, row['Qty']).font = regular_font
            ws.cell(r, 5).number_format = "#,##0"
            ws.cell(r, 6, sanitize_excel_value(row['Unit'])).font = regular_font
            ws.cell(r, 7, row['Cost']).font = regular_font
            ws.cell(r, 7).number_format = NUM_FMT_CURRENCY
            ws.cell(r, 8, row['Total']).font = bold_font
            ws.cell(r, 8).number_format = NUM_FMT_CURRENCY
            ws.cell(r, 9, sanitize_excel_value(row['Requested_By'])).font = regular_font

            for col_i in range(1, 10):
                ws.cell(r, col_i).border = thin_border
            r += 1

        # Log Subtotal
        ws.cell(r, 4, "TOTAL ISSUED").font = bold_font
        ws.cell(r, 5, f"=SUM(E{start_log_r}:E{r-1})").font = bold_font
        ws.cell(r, 5).number_format = "#,##0"
        ws.cell(r, 8, f"=SUM(H{start_log_r}:H{r-1})").font = bold_font
        ws.cell(r, 8).number_format = NUM_FMT_CURRENCY
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
        ws_all.cell(r_idx, 2, sanitize_excel_value(row['SR_Number'])).font = regular_font
        ws_all.cell(r_idx, 3, sanitize_excel_value(row['Category'])).font = bold_font
        ws_all.cell(r_idx, 4, sanitize_excel_value(row['Item_Code'])).font = regular_font
        ws_all.cell(r_idx, 5, sanitize_excel_value(row['Item_Name'])).font = bold_font
        ws_all.cell(r_idx, 6, row['Qty']).font = regular_font
        ws_all.cell(r_idx, 6).number_format = "#,##0"
        ws_all.cell(r_idx, 7, sanitize_excel_value(row['Unit'])).font = regular_font
        ws_all.cell(r_idx, 8, row['Cost']).font = regular_font
        ws_all.cell(r_idx, 8).number_format = NUM_FMT_CURRENCY
        ws_all.cell(r_idx, 9, row['Total']).font = bold_font
        ws_all.cell(r_idx, 9).number_format = NUM_FMT_CURRENCY
        ws_all.cell(r_idx, 10, sanitize_excel_value(row['Requested_By'])).font = regular_font
        ws_all.cell(r_idx, 11, sanitize_excel_value(row.get('Source_File', ''))).font = regular_font

        for col_i in range(1, 12):
            ws_all.cell(r_idx, col_i).border = thin_border

    if room_cost_daily_df is not None and not room_cost_daily_df.empty:
        ws_cost = wb.create_sheet(title="Room Cost Daily")
        ws_cost.views.sheetView[0].showGridLines = True
        ws_cost["A1"] = "Daily Room Cost"
        ws_cost["A1"].font = title_font
        ws_cost["A2"] = f"{sanitize_excel_value(report_title)} | Current MTD SR rate applied retroactively"
        ws_cost["A2"].font = subtitle_font

        monthly_headers = [
            "Month", "Average Occupancy (%)", "Room Nights", "Water Total (IDR)", "Electricity Total (IDR)",
            "Linen Laundry Total (IDR)", "SR Total (IDR)", "Total Room Cost (IDR)",
            "Total Cost / Room (IDR)",
        ]
        for col_idx, header in enumerate(monthly_headers, 1):
            cell = ws_cost.cell(4, col_idx, header)
            cell.font = white_bold
            cell.fill = dark_slate_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        monthly_row = room_cost_monthly_df.iloc[0] if room_cost_monthly_df is not None and not room_cost_monthly_df.empty else None
        if monthly_row is not None:
            monthly_values = [
                monthly_row["Month"], monthly_row["Occupancy_Average"], monthly_row["Room_Nights"], monthly_row["Water_Total"],
                monthly_row["Electricity_Total"], monthly_row["Linen_Laundry_Total"],
                monthly_row["SR_Total"], monthly_row["Total_Room_Cost"], monthly_row["Total_Cost_Per_Room"],
            ]
            for col_idx, value in enumerate(monthly_values, 1):
                cell = ws_cost.cell(5, col_idx, None if pd.isna(value) else value)
                cell.font = bold_font
                cell.border = thin_border
                if col_idx == 2:
                    cell.number_format = "0.00%"
                    if not pd.isna(value):
                        cell.value = value / 100
                elif col_idx >= 4:
                    cell.number_format = NUM_FMT_CURRENCY

        daily_headers = [
            "Date", "Occupancy (%)", "Room Nights", "Water Total (IDR)", "Water / Room (IDR)",
            "Electricity Total (IDR)", "Electricity / Room (IDR)",
            "Linen Laundry Total (IDR)", "Linen Laundry / Room (IDR)",
            "SR Received (IDR)", "SR MTD Spend (IDR)", "MTD Room Nights",
            "SR Cost / Room (IDR)", "Total Cost / Room (IDR)", "Total Room Cost (IDR)",
        ]
        for col_idx, header in enumerate(daily_headers, 1):
            cell = ws_cost.cell(8, col_idx, header)
            cell.font = white_bold
            cell.fill = blue_header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        daily_columns = [
            "Date", "Occupancy", "Room_Nights", "Water", "Water_Per_Room", "Electricity",
            "Electricity_Per_Room", "Linen_Laundry", "Linen_Laundry_Per_Room",
            "SR_Received_On_Date", "MTD_SR_Spend", "MTD_Room_Nights", "SR_Per_Room",
            "Total_Cost_Per_Room", "Total_Room_Cost",
        ]
        for row_idx, (_, row) in enumerate(room_cost_daily_df.iterrows(), 9):
            for col_idx, column in enumerate(daily_columns, 1):
                value = row[column]
                cell = ws_cost.cell(row_idx + 1, col_idx, None if pd.isna(value) else value)
                cell.border = thin_border
                cell.font = regular_font
                if column == "Occupancy" and not pd.isna(value):
                    cell.value = value / 100
                    cell.number_format = "0.00%"
                elif column not in {"Date", "Occupancy", "Room_Nights", "MTD_Room_Nights"}:
                    cell.number_format = NUM_FMT_CURRENCY

    # Auto-adjust column widths across all sheets

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


def create_morning_briefing_template(
    month: str,
    sr_data: Optional[pd.DataFrame] = None,
    daily_cost_data: Optional[pd.DataFrame] = None,
) -> bytes:
    """Build an eight-sheet briefing workbook, prefilling available SR and daily data."""
    try:
        month_start = datetime.strptime(month, "%Y-%m").date()
    except (TypeError, ValueError) as exc:
        raise ValueError("Month must use YYYY-MM format.") from exc

    from calendar import monthrange

    days_in_month = monthrange(month_start.year, month_start.month)[1]
    month_sr_data = pd.DataFrame()
    if sr_data is not None and not sr_data.empty and {"Date", "Total"}.issubset(sr_data.columns):
        sr_dates = pd.to_datetime(sr_data["Date"], errors="coerce", format="mixed")
        month_sr_data = sr_data.loc[
            sr_dates.dt.strftime("%Y-%m") == month,
            [column for column in ("Date", "SR_Number", "Category", "Total") if column in sr_data.columns],
        ].copy()
    sr_rows = max(300, len(month_sr_data))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    navy_fill = PatternFill(fill_type="solid", fgColor="17324D")
    teal_fill = PatternFill(fill_type="solid", fgColor="167C80")
    light_fill = PatternFill(fill_type="solid", fgColor="EAF3F3")
    input_fill = PatternFill(fill_type="solid", fgColor="FFF4D6")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="17324D")
    note_font = Font(name="Calibri", size=10, italic=True, color="526579")
    body_font = Font(name="Calibri", size=10, color="243746")
    thin_border = Border(
        left=Side(style="thin", color="D5E0E5"),
        right=Side(style="thin", color="D5E0E5"),
        top=Side(style="thin", color="D5E0E5"),
        bottom=Side(style="thin", color="D5E0E5"),
    )

    sheet_specs = {
        "Briefing": ["Morning Briefing", "Value", "Room Attendant", "Section / Floor", "Credit / Amount", "Description", "Status", "PIC", "Notes", "Follow Up"],
        "SR Total MTD": ["Date", "SR Number", "Category", "Total (IDR)", "MTD Total (IDR)"],
        "Daily Linen Cost MTD": ["Date", "Daily Linen Cost (IDR)", "MTD Total (IDR)"],
        "Bon Vivo Guest Laundry MTD": [
            "Date", "Room Nights", "Guest Laundry Cost (IDR)",
            "Cost per Room (IDR)", "MTD Total (IDR)",
        ],
        "Project Public Area": [
            "Date", "Project / Area", "Description", "Qty", "Unit", "Unit Cost (IDR)",
            "Total (IDR)", "Status", "PIC", "Notes",
        ],
        "Daily Public Area": [
            "Date", "Area", "Description", "Qty", "Unit Cost (IDR)",
            "Total (IDR)", "PIC", "Notes",
        ],
        "Project RA": [
            "Date", "Room Attendant", "Section / Floor", "Credit Kamar", "Project / Activity", "Notes",
        ],
        "Room Cost MTD": [
            "Date", "Occupancy (%)", "Room Nights", "SR Supplies (IDR)",
            "Air (IDR)", "Listrik (IDR)", "Linen (IDR)",
            "Total Room Cost (IDR)", "Room Cost per Room (IDR)",
        ],
    }

    def create_sheet(name: str, headers: list[str], note: str, rows: int) -> openpyxl.worksheet.worksheet.Worksheet:
        worksheet = wb.create_sheet(name)
        worksheet.freeze_panes = "A5"
        worksheet.sheet_view.showGridLines = False
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        worksheet.cell(1, 1, name).font = title_font
        worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
        worksheet.cell(2, 1, f"Morning Briefing | {month} | {note}").font = note_font
        for column, header in enumerate(headers, 1):
            cell = worksheet.cell(4, column, header)
            cell.font = header_font
            cell.fill = navy_fill if name == "Briefing" else teal_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        worksheet.row_dimensions[4].height = 32
        for row_idx in range(5, 5 + rows):
            for col_idx in range(1, len(headers) + 1):
                cell = worksheet.cell(row_idx, col_idx)
                cell.font = body_font
                cell.border = thin_border
                if row_idx % 2 == 0:
                    cell.fill = light_fill
        for col_idx, header in enumerate(headers, 1):
            worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(len(header) + 3, 15), 34)
        return worksheet

    briefing_note = "Ringkasan satu hari; ubah tanggal untuk melihat hari lain. Rincian MTD ada di sheet pendukung."
    briefing = create_sheet("Briefing", sheet_specs["Briefing"], briefing_note, 112)
    sr_sheet = create_sheet(
        "SR Total MTD", sheet_specs["SR Total MTD"],
        "Masukkan satu baris per voucher atau item SR; total MTD dihitung otomatis.", sr_rows,
    )
    linen = create_sheet(
        "Daily Linen Cost MTD", sheet_specs["Daily Linen Cost MTD"],
        "Masukkan biaya linen harian; total MTD dihitung otomatis.", days_in_month,
    )
    laundry = create_sheet(
        "Bon Vivo Guest Laundry MTD", sheet_specs["Bon Vivo Guest Laundry MTD"],
        "Masukkan biaya guest laundry harian; biaya per room dan total MTD dihitung otomatis.", days_in_month,
    )
    project_pa = create_sheet(
        "Project Public Area", sheet_specs["Project Public Area"],
        "Masukkan rincian biaya proyek Public Area.", 200,
    )
    daily_pa = create_sheet(
        "Daily Public Area", sheet_specs["Daily Public Area"],
        "Masukkan biaya operasional Public Area harian.", 200,
    )
    project_ra = create_sheet(
        "Project RA", sheet_specs["Project RA"],
        "Masukkan satu baris per RA per tanggal; isi credit kamar dan project/activity untuk Briefing.", 200,
    )
    room_cost = create_sheet(
        "Room Cost MTD", sheet_specs["Room Cost MTD"],
        "Isi occupancy, room nights, air, listrik, dan linen; SR dan total dihitung otomatis.", days_in_month,
    )

    first_date = month_start
    for day in range(days_in_month):
        row = 5 + day
        current_date = first_date.replace(day=day + 1)
        for worksheet in (linen, laundry, room_cost):
            worksheet.cell(row, 1, current_date)
            worksheet.cell(row, 1).number_format = "dd mmm yyyy"
        for col_idx in (2, 3, 5):
            linen.cell(row, col_idx).fill = input_fill
        for col_idx in (2, 3, 5, 6, 7):
            room_cost.cell(row, col_idx).fill = input_fill
        laundry.cell(row, 2).fill = input_fill
        laundry.cell(row, 3).fill = input_fill
        laundry.cell(row, 4, f'=IFERROR(C{row}/B{row},"")')
        laundry.cell(row, 5, f'=SUM($C$5:C{row})')
        linen.cell(row, 2).fill = input_fill
        linen.cell(row, 3, f'=SUM($B$5:B{row})')

        room_cost.cell(row, 4, f'=IFERROR(C{row}*SUM(\'SR Total MTD\'!$D$5:$D${4 + sr_rows})/SUM($C$5:$C${4 + days_in_month}),0)')
        room_cost.cell(row, 8, f'=IF(COUNTA(C{row},E{row}:G{row})=0,"",SUM(D{row}:G{row}))')
        room_cost.cell(row, 9, f'=IFERROR(H{row}/C{row},"")')

        for worksheet, columns in ((room_cost, (2,)),):
            for col_idx in columns:
                worksheet.cell(row, col_idx).number_format = "0.00%"
        for worksheet, columns in (
            (room_cost, (4, 5, 6, 7, 8, 9)),
            (linen, (2, 3)),
            (laundry, (3, 4, 5)),
        ):
            for col_idx in columns:
                worksheet.cell(row, col_idx).number_format = NUM_FMT_CURRENCY

    yesterday = datetime.now().date() - timedelta(days=1)
    briefing_date = yesterday if yesterday.strftime("%Y-%m") == month else month_start.replace(day=days_in_month)
    briefing.merge_cells("A1:J1")
    briefing["A1"] = "Morning Briefing"
    briefing["A1"].font = title_font
    briefing.merge_cells("A2:J2")
    briefing["A2"] = f"Daily operational summary | {month} | MTD details are in the supporting sheets"
    briefing["A2"].font = note_font
    briefing["A4"] = "Tanggal Briefing"
    briefing["B4"] = briefing_date
    briefing["B4"].number_format = "dd mmm yyyy"
    briefing["B4"].fill = input_fill
    briefing["A5"] = "Tanggal Operasional (hari ini)"
    briefing["B5"] = "=B4+1"
    briefing["B5"].number_format = "dd mmm yyyy"
    briefing["A6"] = "OCCUPANCY & STAFF"
    briefing["A7"] = "Occupancy Semalam (%)"
    briefing["B7"] = f'=IFERROR(INDEX(\'Room Cost MTD\'!$B$5:$B${4 + days_in_month},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${4 + days_in_month},0)),"")'
    briefing["B7"].number_format = "0.00%"
    briefing["A8"] = "Room Sold"
    briefing["B8"] = f'=IFERROR(INDEX(\'Room Cost MTD\'!$C$5:$C${4 + days_in_month},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${4 + days_in_month},0)),"")'
    briefing["A9"] = "Jumlah Staff RA"
    briefing["B9"] = '=IFERROR(ROWS(UNIQUE(FILTER(\'Project RA\'!$B$5:$B$204,(\'Project RA\'!$A$5:$A$204=$B$4)*(\'Project RA\'!$B$5:$B$204<>""),""))),0)'
    briefing["A10"] = "Staff Incharge"
    briefing["B10"].fill = input_fill

    def section_header(row: int, title: str) -> None:
        briefing.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
        cell = briefing.cell(row, 1, title)
        cell.font = header_font
        cell.fill = teal_fill

    def table_header(row: int, labels: list[str]) -> None:
        for column, label in enumerate(labels, 1):
            cell = briefing.cell(row, column, label)
            cell.font = header_font
            cell.fill = navy_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

    section_header(12, "ROOM CREDIT PER RA | HARI KEMARIN")
    table_header(13, ["Room Attendant", "Section / Floor", "Credit Kamar", "Project / Activity", "Notes"])
    briefing["A14"] = '=IFERROR(TAKE(FILTER(\'Project RA\'!$B$5:$F$204,\'Project RA\'!$A$5:$A$204=$B$4,""),20),"")'

    section_header(36, "ROOM COST | DETAIL HARI KEMARIN")
    table_header(37, ["Komponen", "Biaya per Kamar (IDR)"])
    room_cost_rows = [
        (38, "SR Supplies / kamar", "D"), (39, "Air / kamar", "E"),
        (40, "Listrik / kamar", "F"), (41, "Linen / kamar", "G"),
    ]
    for row, label, column in room_cost_rows:
        briefing.cell(row, 1, label)
        briefing.cell(row, 2, f'=IFERROR(INDEX(\'Room Cost MTD\'!${column}$5:${column}${4 + days_in_month},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${4 + days_in_month},0))/INDEX(\'Room Cost MTD\'!$C$5:$C${4 + days_in_month},MATCH($B$4,\'Room Cost MTD\'!$A$5:$A${4 + days_in_month},0)),"")')
        briefing.cell(row, 2).number_format = NUM_FMT_CURRENCY
    briefing["A42"] = "TOTAL ROOM COST / KAMAR"
    briefing["B42"] = '=IF(COUNT(B38:B41)<4,"",SUM(B38:B41))'
    briefing["B42"].number_format = NUM_FMT_CURRENCY

    section_header(45, "BON VIVO GUEST LAUNDRY | HARI KEMARIN")
    table_header(46, ["Room Nights", "Laundry Cost (IDR)", "Cost / Room (IDR)"])
    for col_idx, source_col in ((1, "B"), (2, "C"), (3, "D")):
        briefing.cell(47, col_idx, f'=IFERROR(INDEX(\'Bon Vivo Guest Laundry MTD\'!${source_col}$5:${source_col}${4 + days_in_month},MATCH($B$4,\'Bon Vivo Guest Laundry MTD\'!$A$5:$A${4 + days_in_month},0)),"")')
        if col_idx > 1:
            briefing.cell(47, col_idx).number_format = NUM_FMT_CURRENCY

    section_header(50, "PROJECT RA | HARI INI")
    table_header(51, ["Room Attendant", "Section / Floor", "Credit Kamar", "Project / Activity", "Notes"])
    briefing["A52"] = '=IFERROR(TAKE(FILTER(\'Project RA\'!$B$5:$F$204,\'Project RA\'!$A$5:$A$204=$B$5,""),15),"")'

    section_header(69, "PROJECT PUBLIC AREA | HARI INI")
    table_header(70, ["Project / Area", "Description", "Qty", "Unit", "Unit Cost", "Total", "Status", "PIC", "Notes"])
    briefing["A71"] = '=IFERROR(TAKE(FILTER(\'Project Public Area\'!$B$5:$J$204,\'Project Public Area\'!$A$5:$A$204=$B$5,""),15),"")'

    section_header(88, "DAILY PUBLIC AREA | HARI INI")
    table_header(89, ["Area", "Description", "Qty", "Unit Cost", "Total", "PIC", "Notes"])
    briefing["A90"] = '=IFERROR(TAKE(FILTER(\'Daily Public Area\'!$B$5:$H$204,\'Daily Public Area\'!$A$5:$A$204=$B$5,""),15),"")'

    section_header(107, "REMARK / FOLLOW UP")
    briefing.merge_cells("A108:J110")
    briefing["A108"].fill = input_fill
    briefing["A108"].alignment = Alignment(vertical="top", wrap_text=True)
    briefing.row_dimensions[108].height = 30
    briefing.row_dimensions[109].height = 24
    briefing.row_dimensions[110].height = 24
    briefing.freeze_panes = "A4"
    briefing.sheet_properties.pageSetUpPr.fitToPage = True
    briefing.page_setup.orientation = "landscape"
    briefing.page_setup.fitToWidth = 1
    briefing.page_setup.fitToHeight = 0
    briefing.print_area = "A1:J110"
    for column in range(1, 11):
        briefing.column_dimensions[get_column_letter(column)].width = [24, 19, 20, 20, 19, 28, 16, 18, 30, 30][column - 1]
    for row in range(1, 111):
        for column in range(1, 11):
            cell = briefing.cell(row, column)
            if cell.value is not None and cell.border == Border():
                cell.border = thin_border

    if daily_cost_data is not None and not daily_cost_data.empty and "Date" in daily_cost_data.columns:
        saved_dates = pd.to_datetime(daily_cost_data["Date"], errors="coerce", format="mixed")
        saved_month_data = daily_cost_data.loc[saved_dates.dt.strftime("%Y-%m") == month]
        date_to_row = {
            month_start.replace(day=day + 1).isoformat(): 5 + day
            for day in range(days_in_month)
        }
        for _, saved_row in saved_month_data.iterrows():
            parsed_date = pd.to_datetime(saved_row["Date"], errors="coerce", format="mixed")
            row_idx = date_to_row.get(parsed_date.date().isoformat())
            if row_idx is None:
                continue
            for column, col_idx in (
                ("Occupancy", 2), ("Room_Nights", 3), ("Water", 5),
                ("Electricity", 6), ("Linen_Laundry", 7),
            ):
                value = saved_row.get(column)
                if pd.notna(value):
                    numeric_value = pd.to_numeric(value, errors="coerce")
                    if pd.notna(numeric_value):
                        room_cost.cell(
                            row_idx, col_idx,
                            float(numeric_value) / 100 if column == "Occupancy" else float(numeric_value),
                        )
                        if column == "Linen_Laundry":
                            linen.cell(row_idx, 2, float(numeric_value))
                            linen.cell(row_idx, 2).number_format = NUM_FMT_CURRENCY

    for row in range(5, 5 + sr_rows):
        sr_sheet.cell(row, 5, f'=SUM($D$5:D{row})')
        sr_sheet.cell(row, 4).fill = input_fill
        for col_idx in (4, 5):
            sr_sheet.cell(row, col_idx).number_format = NUM_FMT_CURRENCY
        for col_idx in (1, 2, 3):
            sr_sheet.cell(row, col_idx).fill = input_fill
    for row_idx, (_, sr_row) in enumerate(month_sr_data.iterrows(), 5):
        parsed_date = pd.to_datetime(sr_row["Date"], errors="coerce", format="mixed")
        sr_sheet.cell(row_idx, 1, parsed_date.to_pydatetime())
        sr_sheet.cell(row_idx, 1).number_format = "dd mmm yyyy"
        for col_idx, column in ((2, "SR_Number"), (3, "Category")):
            if column in sr_row and pd.notna(sr_row[column]):
                sr_sheet.cell(row_idx, col_idx, sanitize_excel_value(sr_row[column]))
        sr_total = pd.to_numeric(sr_row["Total"], errors="coerce")
        sr_sheet.cell(row_idx, 4, None if pd.isna(sr_total) else float(sr_total))
        sr_sheet.cell(row_idx, 4).fill = input_fill
        sr_sheet.cell(row_idx, 4).number_format = NUM_FMT_CURRENCY
    sr_sheet.column_dimensions["A"].width = 17
    project_pa.column_dimensions["C"].width = 32
    daily_pa.column_dimensions["C"].width = 32
    project_ra.column_dimensions["B"].width = 24
    project_ra.column_dimensions["E"].width = 32

    for worksheet in (project_pa, daily_pa):
        for row in range(5, worksheet.max_row + 1):
            qty_col = 4
            cost_col = 6 if worksheet is project_pa else 5
            total_col = 7 if worksheet is project_pa else 6
            worksheet.cell(row, total_col, f'=IFERROR(D{row}*{get_column_letter(cost_col)}{row},"")')
            worksheet.cell(row, total_col).number_format = NUM_FMT_CURRENCY
            for col_idx in range(1, worksheet.max_column + 1):
                if col_idx != total_col:
                    worksheet.cell(row, col_idx).fill = input_fill

    for worksheet in wb.worksheets:
        if worksheet.title != "Briefing":
            worksheet.auto_filter.ref = f"A4:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

