"""Month discovery and reconciliation orchestration."""
import os
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

import parser
import matcher
from ui.security import validate_data_dir


def get_default_data_dir() -> str:
    """Return a portable default data directory for local month folders."""
    env_path = os.environ.get("HOTEL_DATA_DIR")
    if env_path:
        error = validate_data_dir(env_path)
        if error:
            raise ValueError(f"Invalid HOTEL_DATA_DIR: {error}")
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
