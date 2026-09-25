"""Month discovery, demo data, and reconciliation orchestration."""
import os
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

import parser
import matcher


def get_default_data_dir() -> str:
    """Return a portable default data directory for local month folders."""
    env_path = os.environ.get("HOTEL_DATA_DIR")
    if env_path:
        return env_path

    home_dir = os.path.expanduser("~")
    candidates = [
        os.path.join(home_dir, "Documents", "Business Review"),
        os.path.join(home_dir, "Business Review"),
        os.path.join("C:\\", "Users", os.path.basename(home_dir), "Documents", "Business Review"),
        os.path.join("C:\\", "Business Review"),
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return os.path.join(home_dir, "Documents", "Business Review")


def discover_month_folders(base_dir: str) -> Dict[str, str]:
    """Discover month folders under base_dir and return {label: path}."""
    if not base_dir or not os.path.exists(base_dir):
        return {}

    found: Dict[str, str] = {}
    for entry in sorted(os.listdir(base_dir)):
        full_path = os.path.join(base_dir, entry)
        if not os.path.isdir(full_path):
            continue
        label = parser.infer_month_label(entry, default=entry)
        found[label] = full_path
    return found


@st.cache_data(show_spinner=False)
def cached_reconcile_month(
    dtb_path: str,
    is_path: Optional[str],
    cons_path: Optional[str],
    month_label: str,
) -> dict:
    """Parse and reconcile a single month with Streamlit caching."""
    dtb_df = parser.parse_detail_trial_balance(dtb_path)
    is_df = parser.parse_income_statement(is_path) if is_path and os.path.exists(is_path) else pd.DataFrame()
    cons_df = parser.parse_consumption_report(cons_path) if cons_path and os.path.exists(cons_path) else pd.DataFrame()
    return matcher.reconcile_monthly_expenses(dtb_df, is_df, cons_df)


def build_demo_reconciled_data(month_name: str = "Demo August 2026") -> dict:
    """Return a small demo dataset for previewing the dashboard."""
    expense_rows = [
        {"Date": "2026-08-01", "Category": "Guest Supplies", "Item_Name": "Cleo 330ml", "Partner_Vendor": "Cleo", "Qty": 110, "Unit": "pcs", "Unit_Price": 5000, "Amount": 550000, "Voucher_Ref": "DEMO-001", "JRNL": "GL", "Raw_Description": "Cleo mineral water", "Month": month_name},
        {"Date": "2026-08-02", "Category": "Cleaning Supplies", "Item_Name": "Trash Bag 100L", "Partner_Vendor": "Plasticindo", "Qty": 40, "Unit": "pcs", "Unit_Price": 18000, "Amount": 720000, "Voucher_Ref": "DEMO-002", "JRNL": "GL", "Raw_Description": "Garbage bags", "Month": month_name},
        {"Date": "2026-08-03", "Category": "Paper Supplies", "Item_Name": "Toilet Tissue", "Partner_Vendor": "Paper Mart", "Qty": 65, "Unit": "roll", "Unit_Price": 21000, "Amount": 1365000, "Voucher_Ref": "DEMO-003", "JRNL": "GL", "Raw_Description": "Toilet tissue roll", "Month": month_name},
        {"Date": "2026-08-05", "Category": "Media & Utilities", "Item_Name": "Internet Maxindo", "Partner_Vendor": "Maxindo", "Qty": 1, "Unit": "contract", "Unit_Price": 1650000, "Amount": 1650000, "Voucher_Ref": "DEMO-004", "JRNL": "GL", "Raw_Description": "Internet service", "Month": month_name},
        {"Date": "2026-08-07", "Category": "Outsourcing & Laundry", "Item_Name": "Laundry Service", "Partner_Vendor": "Bonvivo", "Qty": 24, "Unit": "lot", "Unit_Price": 145000, "Amount": 3480000, "Voucher_Ref": "DEMO-005", "JRNL": "GL", "Raw_Description": "Laundry outsourcing", "Month": month_name},
        {"Date": "2026-08-09", "Category": "Payroll & SC", "Item_Name": "Daily Worker Allowance", "Partner_Vendor": "HRD", "Qty": 1, "Unit": "payroll", "Unit_Price": 2500000, "Amount": 2500000, "Voucher_Ref": "DEMO-006", "JRNL": "GL", "Raw_Description": "Housekeeping payroll", "Month": month_name},
    ]
    tx_df = pd.DataFrame(expense_rows)
    tx_df['Date'] = pd.to_datetime(tx_df['Date'])
    tx_df['Amount'] = tx_df['Amount'].astype(float)

    category_summary = pd.DataFrame([
        {"Category": "Guest Supplies", "Department": "Housekeeping", "Group": "Operational", "Actual": 550000, "Budget": 600000, "Transaction_Count": 1, "Top_Item": "Cleo 330ml"},
        {"Category": "Cleaning Supplies", "Department": "Housekeeping", "Group": "Operational", "Actual": 720000, "Budget": 700000, "Transaction_Count": 1, "Top_Item": "Trash Bag 100L"},
        {"Category": "Paper Supplies", "Department": "Housekeeping", "Group": "Operational", "Actual": 1365000, "Budget": 1200000, "Transaction_Count": 1, "Top_Item": "Toilet Tissue"},
        {"Category": "Media & Utilities", "Department": "Engineering", "Group": "Utilities", "Actual": 1650000, "Budget": 1500000, "Transaction_Count": 1, "Top_Item": "Internet Maxindo"},
        {"Category": "Outsourcing & Laundry", "Department": "Housekeeping", "Group": "Outsourcing", "Actual": 3480000, "Budget": 3000000, "Transaction_Count": 1, "Top_Item": "Laundry Service"},
        {"Category": "Payroll & SC", "Department": "HR", "Group": "Payroll", "Actual": 2500000, "Budget": 2200000, "Transaction_Count": 1, "Top_Item": "Daily Worker Allowance"},
    ])
    category_summary['Variance_IDR'] = category_summary['Actual'] - category_summary['Budget']
    category_summary['Variance_Pct'] = np.where(category_summary['Budget'] > 0, (category_summary['Variance_IDR'] / category_summary['Budget']) * 100.0, 0.0)
    category_summary['Status'] = np.where(category_summary['Variance_IDR'] > 0, 'Over Budget', 'Under Budget')

    item_summary = tx_df.groupby(['Category', 'Item_Name'], as_index=False).agg(
        Total_Qty=('Qty', 'sum'),
        Unit=('Unit', 'first'),
        Avg_Unit_Price=('Unit_Price', 'mean'),
        Total_Amount=('Amount', 'sum'),
        Order_Count=('Amount', 'count'),
        First_Date=('Date', 'min'),
        Last_Date=('Date', 'max')
    )
    item_summary['Cat_Total'] = item_summary['Total_Amount'].sum()
    item_summary['Pct_Of_Category'] = np.where(item_summary['Cat_Total'] > 0, (item_summary['Total_Amount'] / item_summary['Cat_Total'] * 100.0).round(2), 0.0)

    total_budget = category_summary['Budget'].sum()
    total_spent = tx_df['Amount'].sum()
    metrics = {
        'total_spent': float(total_spent),
        'total_budget': float(total_budget),
        'variance_idr': float(total_spent - total_budget),
        'variance_pct': (float(total_spent - total_budget) / total_budget * 100.0) if total_budget > 0 else 0.0,
        'transaction_count': len(tx_df),
        'category_count': len(category_summary),
        'unique_items_count': len(item_summary),
    }

    return {
        'transactions': tx_df,
        'category_summary': category_summary,
        'top_cost_drivers': tx_df.groupby(['Category', 'Item_Name', 'Partner_Vendor'], as_index=False).agg(
            Total_Amount=('Amount', 'sum'),
            Total_Qty=('Qty', 'sum'),
            Unit=('Unit', 'first'),
            Trx_Count=('Amount', 'count')
        ).sort_values(by='Total_Amount', ascending=False).reset_index(drop=True),
        'item_summary': item_summary,
        'metrics': metrics,
    }


def combine_month_datasets(month_data_dict: dict) -> Tuple[Optional[dict], str]:
    """Combine multiple monthly reconciled datasets into a single dataset."""
    if not month_data_dict:
        return None, ""

    all_trxs = []
    for month_name, m_data in month_data_dict.items():
        trx = m_data['transactions'].copy()
        if 'Month' not in trx.columns or trx['Month'].isna().all():
            trx['Month'] = month_name
        all_trxs.append(trx)

    combined_trxs = pd.concat(all_trxs, ignore_index=True)

    combined_cat_summary = pd.concat([m['category_summary'] for m in month_data_dict.values()], ignore_index=True)
    if not combined_cat_summary.empty:
        grp_cat = combined_cat_summary.groupby('Category', as_index=False).agg({
            'Department': 'first',
            'Group': 'first',
            'Actual': 'sum',
            'Budget': 'sum',
            'Transaction_Count': 'sum',
            'Top_Item': 'first'
        })
        grp_cat['Variance_IDR'] = grp_cat['Actual'] - grp_cat['Budget']
        grp_cat['Variance_Pct'] = np.where(
            grp_cat['Budget'] > 0,
            (grp_cat['Variance_IDR'] / grp_cat['Budget']) * 100.0,
            0.0
        )
        grp_cat['Status'] = np.where(
            grp_cat['Budget'] > 0,
            np.where(grp_cat['Variance_IDR'] > 0, 'Over Budget', 'Under Budget'),
            np.where(grp_cat['Actual'] > 0, 'Unbudgeted', 'On Track')
        )
        grp_cat = grp_cat.sort_values(by='Actual', ascending=False).reset_index(drop=True)
    else:
        grp_cat = pd.DataFrame()

    top_items_df = (
        combined_trxs.groupby(['Category', 'Item_Name', 'Partner_Vendor'], as_index=False)
        .agg(
            Total_Amount=('Amount', 'sum'),
            Total_Qty=('Qty', 'sum'),
            Unit=('Unit', 'first'),
            Trx_Count=('Amount', 'count')
        )
        .sort_values(by='Total_Amount', ascending=False)
        .reset_index(drop=True)
    )

    combined_trxs_copy = combined_trxs.copy()
    cat_totals = combined_trxs_copy.groupby('Category')['Amount'].transform('sum')
    combined_trxs_copy['Cat_Total'] = cat_totals

    item_summary_df = combined_trxs_copy.groupby(['Category', 'Item_Name'], as_index=False).agg(
        Total_Qty=('Qty', 'sum'),
        Unit=('Unit', lambda u: next((str(x) for x in u if str(x).strip() and str(x) != 'nan'), '')),
        Avg_Unit_Price=('Unit_Price', 'mean'),
        Total_Amount=('Amount', 'sum'),
        Order_Count=('Amount', 'count'),
        First_Date=('Date', 'min'),
        Last_Date=('Date', 'max'),
        Cat_Total=('Cat_Total', 'first')
    )

    item_summary_df['Pct_Of_Category'] = np.where(
        item_summary_df['Cat_Total'] > 0,
        (item_summary_df['Total_Amount'] / item_summary_df['Cat_Total'] * 100.0).round(2),
        0.0
    )
    item_summary_df = item_summary_df.sort_values(
        by=['Category', 'Total_Amount'], ascending=[True, False]
    ).reset_index(drop=True)

    total_spent = combined_trxs['Amount'].sum()
    total_budget = grp_cat['Budget'].sum() if not grp_cat.empty else 0.0
    total_var = total_spent - total_budget
    total_var_pct = (total_var / total_budget * 100.0) if total_budget > 0 else 0.0

    combined_data = {
        'transactions': combined_trxs,
        'category_summary': grp_cat,
        'top_cost_drivers': top_items_df,
        'item_summary': item_summary_df,
        'metrics': {
            'total_spent': total_spent,
            'total_budget': total_budget,
            'variance_idr': total_var,
            'variance_pct': total_var_pct,
            'transaction_count': len(combined_trxs),
            'category_count': len(grp_cat),
            'unique_items_count': len(item_summary_df)
        }
    }

    month_names_str = ", ".join(month_data_dict.keys())
    return combined_data, month_names_str
