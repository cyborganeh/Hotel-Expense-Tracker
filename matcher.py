"""
matcher.py - Reconciliation and enrichment engine for Hotel Santika Depok expenses.
Joins General Ledger (DTB) transactions with Store Consumption lines and Vendor Invoices,
matching 100% of spending back to items, vendors, and budget lines.
"""

from typing import Dict, Any, Optional, Tuple, List
import re
import pandas as pd
import numpy as np


def reconcile_monthly_expenses(
    dtb_df: pd.DataFrame,
    is_df: pd.DataFrame,
    cons_df: pd.DataFrame,
    laundry_data: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Performs full reconciliation and enrichment between DTB, Income Statement, and Consumption reports.
    """
    # Guard: an empty DTB (e.g. no 035* accounts matched) must yield a well-formed
    # empty result instead of crashing downstream consumers.
    if dtb_df is None or len(dtb_df) == 0:
        return {
            'transactions': pd.DataFrame(columns=[
                'Date', 'Category', 'Account_Code', 'Account_Name', 'Item_Name',
                'Partner_Vendor', 'Voucher_Ref', 'Source', 'JRNL', 'Qty', 'Unit',
                'Unit_Price', 'Amount', 'Raw_Description'
            ]),
            'category_summary': pd.DataFrame(columns=[
                'Category', 'Department', 'Group', 'Actual', 'Budget',
                'Variance_IDR', 'Variance_Pct', 'Ratio_Pct', 'Transaction_Count',
                'Status', 'Top_Item'
            ]),
            'top_cost_drivers': pd.DataFrame(columns=[
                'Category', 'Item_Name', 'Partner_Vendor', 'Total_Amount',
                'Total_Qty', 'Unit', 'Trx_Count'
            ]),
            'item_summary': pd.DataFrame(columns=[
                'Category', 'Item_Name', 'Total_Qty', 'Unit', 'Avg_Unit_Price',
                'Total_Amount', 'Order_Count', 'First_Date', 'Last_Date',
                'Cat_Total', 'Pct_Of_Category'
            ]),
            'metrics': {
                'total_spent': 0.0, 'total_budget': 0.0, 'variance_idr': 0.0,
                'variance_pct': 0.0, 'transaction_count': 0,
                'category_count': 0, 'unique_items_count': 0
            }
        }

    # 1. Enrich DTB transactions with item details from Consumption Report
    enriched_rows = []

    # Prepare a consumable copy of cons_df to match items by voucher and amount
    unmatched_cons = cons_df.copy()
    unmatched_cons['used'] = False

    for _, r in dtb_df.iterrows():
        ref = str(r['Ref']).strip() if pd.notna(r['Ref']) else ''
        amt = float(r['Net_Amount'])
        cat = str(r['Category'])
        partner = str(r['Partner']).strip() if pd.notna(r['Partner']) and r['Partner'] != 'None' else ''
        desc = str(r['Description']).strip() if pd.notna(r['Description']) and r['Description'] != 'None' else ''
        jrnl = str(r['JRNL']).strip() if pd.notna(r['JRNL']) else ''

        item_name = ''
        qty = 1.0
        unit = ''
        unit_price = amt

        # Try to match store vouchers (GDC/OUT or HK/IN)
        if ('GDC/OUT' in ref or 'HK/IN' in ref) and not unmatched_cons.empty:
            # Look for exact voucher and amount match
            candidates = unmatched_cons[
                (~unmatched_cons['used']) &
                (unmatched_cons['Voucher_No'] == ref) &
                (np.isclose(unmatched_cons['Amount'], amt, atol=1.0))
            ]

            if not candidates.empty:
                match_idx = candidates.index[0]
                matched_item = candidates.loc[match_idx]
                unmatched_cons.at[match_idx, 'used'] = True

                item_name = str(matched_item['Item_Name'])
                qty = float(matched_item['Qty'])
                unit = str(matched_item['Unit'])
                unit_price = float(matched_item['Unit_Price'])
            else:
                # Fallback: match by voucher only if single remaining
                voucher_candidates = unmatched_cons[
                    (~unmatched_cons['used']) &
                    (unmatched_cons['Voucher_No'] == ref)
                ]
                if len(voucher_candidates) == 1:
                    match_idx = voucher_candidates.index[0]
                    matched_item = voucher_candidates.loc[match_idx]
                    unmatched_cons.at[match_idx, 'used'] = True

                    item_name = str(matched_item['Item_Name'])
                    qty = float(matched_item['Qty'])
                    unit = str(matched_item['Unit'])
                    unit_price = float(matched_item['Unit_Price'])

        # If not matched via store voucher, determine meaningful item/vendor description
        if not item_name:
            if partner and partner.lower() != 'none':
                item_name = f"{partner} - {desc}" if desc else partner
            elif desc and desc.lower() != 'none':
                item_name = desc
            else:
                item_name = cat

        # Format partner name if empty
        if not partner:
            if 'pjt' in desc.lower() or 'cleaning' in desc.lower():
                partner = 'PT. Pandu Jasa Terpadu'
            elif 'bonvivo' in desc.lower():
                partner = 'Bona Natu Cemerlang Laundry (Bonvivo)'
            elif 'drop n go' in desc.lower():
                partner = 'Drop n Go Laundry'
            elif 'maxindo' in desc.lower() or 'internet' in desc.lower():
                partner = 'Maxindo Mitra Solusi, PT'
            elif 'mnc' in desc.lower() or 'indovision' in desc.lower():
                partner = 'MNC Sky Vision, PT'
            elif 'telkom' in desc.lower():
                partner = 'Telkom'
            elif 'cleo' in item_name.lower():
                partner = 'Store / Cleo Water'
            elif 'payroll' in desc.lower() or 'salary' in desc.lower() or 'gaji' in desc.lower():
                partner = 'Internal Payroll'
            elif 'service charge' in desc.lower() or 'sc ' in desc.lower():
                partner = 'Service Charge Distribution'
            else:
                partner = 'Internal / Store'

        enriched_rows.append({
            'Date': r['Date'],
            'Category': cat,
            'Account_Code': r['Account_Code'],
            'Account_Name': r['Account_Name'],
            'Item_Name': item_name,
            'Partner_Vendor': partner,
            'Voucher_Ref': ref,
            'Source': r['Source'],
            'JRNL': jrnl,
            'Qty': qty,
            'Unit': unit,
            'Unit_Price': unit_price,
            'Amount': amt,
            'Raw_Description': desc
        })

    enriched_df = pd.DataFrame(enriched_rows)

    # 2. Build Category Summary matching with Income Statement Budget & Actual
    cat_summary_rows = []
    is_clean = is_df[~is_df['Is_Subtotal']].copy() if not is_df.empty and 'Is_Subtotal' in is_df.columns else pd.DataFrame()
    
    # Filter out subtotal rows from Income Statement for clean comparison
    is_clean = is_df[~is_df['Is_Subtotal']].copy() if not is_df.empty else pd.DataFrame()

    # Map categories to IS rows
    for cat_name, group_df in enriched_df.groupby('Category'):
        actual_spent = group_df['Amount'].sum()
        trx_count = len(group_df)
        
        # Match with IS row
        budget = 0.0
        is_actual = actual_spent
        is_ratio = 0.0
        dept_name = 'Room Division'
        group_type = 'Other Expenses'

        if not is_clean.empty:
            match = is_clean[is_clean['Description'].str.strip().str.lower() == cat_name.strip().lower()]
            if match.empty and cat_name.strip():
                # Fuzzy or partial match (escape so category names are treated literally)
                match = is_clean[is_clean['Description'].str.contains(re.escape(cat_name[:8]), case=False, na=False)]
            
            if not match.empty:
                m_row = match.iloc[0]
                budget = float(m_row['Budget'])
                is_actual = float(m_row['Actual'])
                is_ratio = float(m_row['Ratio_Pct'])
                dept_name = str(m_row['Department'])
                group_type = str(m_row['Group'])

        var_idr = actual_spent - budget
        var_pct = (var_idr / budget * 100.0) if budget > 0 else 0.0

        status = 'On Track'
        if budget > 0:
            if var_idr > 0:
                status = 'Over Budget'
            else:
                status = 'Under Budget'
        elif actual_spent > 0 and budget == 0:
            status = 'Unbudgeted'

        # Find top item in this category
        top_item = group_df.groupby('Item_Name')['Amount'].sum().sort_values(ascending=False).index[0]

        cat_summary_rows.append({
            'Category': cat_name,
            'Department': dept_name,
            'Group': group_type,
            'Actual': actual_spent,
            'Budget': budget,
            'Variance_IDR': var_idr,
            'Variance_Pct': var_pct,
            'Ratio_Pct': is_ratio,
            'Transaction_Count': trx_count,
            'Status': status,
            'Top_Item': top_item
        })

    cat_summary_df = pd.DataFrame(cat_summary_rows)
    if not cat_summary_df.empty:
        cat_summary_df = cat_summary_df.sort_values(by='Actual', ascending=False).reset_index(drop=True)

    # 3. Top Cost Drivers across the hotel
    top_items_df = (
        enriched_df.groupby(['Category', 'Item_Name', 'Partner_Vendor'])
        .agg(
            Total_Amount=('Amount', 'sum'),
            Total_Qty=('Qty', 'sum'),
            Unit=('Unit', 'first'),
            Trx_Count=('Amount', 'count')
        )
        .reset_index()
        .sort_values(by='Total_Amount', ascending=False)
        .reset_index(drop=True)
    )

    # 4. Item-level summary with total quantities, unit costs, and category share
    enriched_df_copy = enriched_df.copy()
    cat_totals = enriched_df_copy.groupby('Category')['Amount'].transform('sum')
    enriched_df_copy['Cat_Total'] = cat_totals

    item_summary_df = enriched_df_copy.groupby(['Category', 'Item_Name']).agg(
        Total_Qty=('Qty', 'sum'),
        Unit=('Unit', lambda u: next((x for x in u if x and str(x).strip()), '')),
        Avg_Unit_Price=('Unit_Price', 'mean'),
        Total_Amount=('Amount', 'sum'),
        Order_Count=('Amount', 'count'),
        First_Date=('Date', 'min'),
        Last_Date=('Date', 'max'),
        Cat_Total=('Cat_Total', 'first')
    ).reset_index()

    item_summary_df['Pct_Of_Category'] = (
        item_summary_df['Total_Amount'] / item_summary_df['Cat_Total'] * 100.0
    ).round(2)
    item_summary_df = item_summary_df.sort_values(
        by=['Category', 'Total_Amount'], ascending=[True, False]
    ).reset_index(drop=True)

    # 5. Total metrics
    total_spent = enriched_df['Amount'].sum()
    total_budget = cat_summary_df['Budget'].sum() if not cat_summary_df.empty else 0.0
    total_var = total_spent - total_budget
    total_var_pct = (total_var / total_budget * 100.0) if total_budget > 0 else 0.0

    return {
        'transactions': enriched_df,
        'category_summary': cat_summary_df,
        'top_cost_drivers': top_items_df,
        'item_summary': item_summary_df,
        'metrics': {
            'total_spent': total_spent,
            'total_budget': total_budget,
            'variance_idr': total_var,
            'variance_pct': total_var_pct,
            'transaction_count': len(enriched_rows),
            'category_count': len(cat_summary_rows),
            'unique_items_count': len(item_summary_df)
        }
    }
