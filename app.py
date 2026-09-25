"""
app.py - Streamlit Web Application for Hotel Santika Depok
Interactive Expense Separator & Spending Tracker.
"""

import os
import io
import glob
import re
from typing import Tuple, Optional, Dict, Any, List
import numpy as np
import streamlit as st
import pandas as pd
import plotly.express as px  # noqa: F401  (re-imported so ui_components picks up the themed template)
import plotly.graph_objects as go  # noqa: F401  (kept for parity with ui_components)

import parser
import matcher
import exporter
import sr_parser
import ui_components
import importlib
importlib.reload(parser)
importlib.reload(matcher)
importlib.reload(exporter)
importlib.reload(sr_parser)
importlib.reload(ui_components)


def get_default_data_dir() -> str:
    """Return a portable default data directory for local month folders."""
    env_path = os.environ.get("HOTEL_DATA_DIR")
    if env_path:
        return env_path

    candidates = []
    home_dir = os.path.expanduser("~")
    candidates.extend([
        os.path.join(home_dir, "Documents", "Business Review"),
        os.path.join(home_dir, "Business Review"),
        os.path.join("C:\\", "Users", os.path.basename(home_dir), "Documents", "Business Review"),
        os.path.join("C:\\", "Business Review"),
    ])

    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate

    return os.path.join(home_dir, "Documents", "Business Review")


# Configure Streamlit page
st.set_page_config(
    page_title="Hotel Spending Tracker & Separator",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Theme State (defaults to Dark Mode as requested)
if "theme_mode" not in st.session_state:
    st.session_state["theme_mode"] = "dark"

# Sidebar Branding & Theme Mode Choice
st.sidebar.image("https://img.icons8.com/color/96/hotel-star.png", width=64)
st.sidebar.title("Santika Depok")
st.sidebar.markdown("**Room Division & Housekeeping**")

theme_choice = st.sidebar.radio(
    "Theme Appearance",
    ["🌙 Dark Mode", "☀️ Light Mode"],
    index=0 if st.session_state["theme_mode"] == "dark" else 1,
    horizontal=True,
    key="theme_mode_selector"
)
st.session_state["theme_mode"] = "dark" if "Dark" in theme_choice else "light"
is_dark = st.session_state["theme_mode"] == "dark"

# Apply theme CSS variables based on session state
# The consolidated CSS uses [data-theme] attribute on <html> for light/dark switching
theme_mode = st.session_state.get("theme_mode", "dark")
st.markdown(f"""
<script>
    document.documentElement.setAttribute('data-theme', '{theme_mode}');
</script>
""", unsafe_allow_html=True)

# Theme CSS variables (dark defaults + light overrides)
st.markdown("""
<style>
    /* Theme root variables — dark mode defaults */
    :root {
        --bg-main: #0B0F19;
        --bg-card: #151C2C;
        --bg-card-hover: #1A2338;
        --bg-sidebar: #080D1A;
        --text-main: #F8FAFC;
        --text-muted: #94A3B8;
        --border-subtle: rgba(255, 255, 255, 0.08);
        --border-hover: rgba(59, 130, 246, 0.5);
        --accent: #3B82F6;
        --delta-pos: #34D399;
        --delta-neg: #F87171;
        --card-shadow: 0 4px 16px rgba(0, 0, 0, 0.35);
        --card-shadow-hover: 0 6px 22px rgba(0, 0, 0, 0.5);
    }

    /* Light mode overrides — applied when data-theme="light" */
    [data-theme="light"] {
        --bg-main: #F8FAFC;
        --bg-card: #FFFFFF;
        --bg-card-hover: #F1F5F9;
        --bg-sidebar: #FFFFFF;
        --text-main: #0F172A;
        --text-muted: #475569;
        --border-subtle: #E2E8F0;
        --border-hover: rgba(37, 99, 235, 0.45);
        --accent: #2563EB;
        --delta-pos: #15803D;
        --delta-neg: #DC2626;
        --card-shadow: 0 1px 4px rgba(0, 0, 0, 0.05);
        --card-shadow-hover: 0 4px 14px rgba(0, 0, 0, 0.08);
    }
</style>
""", unsafe_allow_html=True)

# Inject shared metric-card and UI CSS from ui_components
st.markdown(ui_components.get_metric_card_css(), unsafe_allow_html=True)

# Helper formatting
def format_idr(val):
    if pd.isna(val):
        return "Rp 0"
    return f"Rp {val:,.0f}"

def format_pct(val):
    if pd.isna(val):
        return "0.0%"
    return f"{val:+.1f}%"


def show_validation_banner(messages: List[str], level: str = "warning"):
    if not messages:
        return
    if level == "error":
        st.error("### Data validation issues")
    else:
        st.warning("### Data validation issues")
    for message in messages:
        st.caption(f"• {message}")


@st.cache_data(show_spinner=False)
def cached_reconcile_month(dtb_path: str, is_path: Optional[str], cons_path: Optional[str], month_label: str) -> dict:
    dtb_df = parser.parse_detail_trial_balance(dtb_path)
    is_df = parser.parse_income_statement(is_path) if is_path and os.path.exists(is_path) else pd.DataFrame()
    cons_df = parser.parse_consumption_report(cons_path) if cons_path and os.path.exists(cons_path) else pd.DataFrame()
    return matcher.reconcile_monthly_expenses(dtb_df, is_df, cons_df)


def discover_month_folders(base_dir: str) -> Dict[str, str]:
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


def build_demo_reconciled_data(month_name: str = "Demo August 2026") -> dict:
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


def render_sr_separator():
    """
    Warehouse Stock Request (SR) Separator UI.
    Extracts, categorizes, and analyzes items from Gudang Central SR PDFs.
    """
    base_review_dir = get_default_data_dir()
    BASE_SR_DIR = os.environ.get("HOTEL_SR_DIR") or os.path.join(base_review_dir, "SR REPORT")
    has_local_sr = os.path.exists(BASE_SR_DIR)

    st.sidebar.markdown("### 📦 SR Data Source")
    sr_options = []
    if has_local_sr:
        sr_options.append("Load Local September SRs")
    sr_options.append("Upload SR PDFs")

    sr_source = st.sidebar.radio("Select Source", sr_options, index=0)

    sr_df = pd.DataFrame()
    sr_report_title = "SR Report September 2026"

    if sr_source == "Load Local September SRs":
        st.sidebar.markdown("**Local Folder:** `SR SEPTEMBER`")
        local_pdfs = sorted(glob.glob(os.path.join(BASE_SR_DIR, "*.pdf")))
        st.sidebar.caption(f"Found {len(local_pdfs)} PDF files:")
        for lp in local_pdfs:
            st.sidebar.caption(f"📄 {os.path.basename(lp)}")
        if local_pdfs:
            with st.spinner("Parsing local SR PDF files..."):
                sr_df = sr_parser.parse_multiple_sr_pdfs(local_pdfs)
            sr_report_title = "SR Report September 2026"

    elif sr_source == "Upload SR PDFs":
        uploaded_pdfs = st.sidebar.file_uploader(
            "Upload Stock Request PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            help="Upload one or multiple SR PDF files from Gudang Central."
        )
        sr_label = st.sidebar.text_input("Report Title / Month", "Stock Request Report")
        sr_report_title = sr_label
        if uploaded_pdfs:
            invalid_names = [
                uploaded.name for uploaded in uploaded_pdfs
                if ".." in uploaded.name or "/" in uploaded.name or "\\" in uploaded.name
            ]
            if invalid_names:
                show_validation_banner([
                    f"Invalid SR filename: {name}. Filenames cannot contain path separators or '..'."
                    for name in invalid_names
                ], level="error")
            else:
                with st.spinner(f"Parsing {len(uploaded_pdfs)} SR PDF(s)..."):
                    sr_df = sr_parser.parse_multiple_sr_pdfs(uploaded_pdfs)

    if sr_df.empty:
        st.info("👈 Please select or upload Stock Request (SR) PDF files from the sidebar to begin.")
        st.markdown("""
        ### About Warehouse Stock Request (SR) Separator
        This tool extracts items from Gudang Central Stock Request PDFs and separates them automatically into:
        - 🛎️ **Guest Supplies**: Soap, slipper, toothbrush, shower cap, coffee, tea, sugar, Cleo water, etc.
        - 🧹 **Cleaning Supplies**: Garbage bags (Plastik sampah), cleaning chemicals, glass cleaner, etc.
        - 🧻 **Paper Supplies**: Facial tissue, hand towel, toilet roll, plastic roll.
        - 📑 **Print & Stationery**: Form blanks, staplers, tape, stationery items.

        **Supported formats**: Stock Request Consumption PDFs from Hotel Santika Depok (`.pdf`).
        """)
        return

    # Generate Excel export buffer
    excel_sr_bytes = exporter.create_sr_separated_excel(sr_df, sr_report_title)

    st.sidebar.divider()
    st.sidebar.markdown("### 📥 Quick Export")
    st.sidebar.download_button(
        label="Download Separated SR Excel (.xlsx)",
        data=excel_sr_bytes,
        file_name=f"Separated_SR_{sr_report_title.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        type="primary"
    )

    # Main Page UI
    st.markdown('<div class="main-header">Warehouse Stock Request (SR) Separator</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Hotel Santika Depok • Gudang Central Issuing to Housekeeping • <b>{sr_report_title}</b></div>', unsafe_allow_html=True)

    # Top KPI Metrics
    tot_amt = sr_df['Total'].sum()
    tot_qty = sr_df['Qty'].sum()
    tot_items = sr_df['Item_Name'].nunique()
    tot_orders = sr_df['SR_Number'].nunique()

    ui_components.render_kpi_row([
        {'label': 'TOTAL SPEND (SR)', 'value': format_idr(tot_amt),
         'delta': f"{len(sr_df)} line items issued", 'delta_tone': 'neutral', 'show_arrow': True},
        {'label': 'TOTAL QUANTITY', 'value': f"{tot_qty:,.0f}",
         'delta': 'Across all units & categories', 'delta_tone': 'pos', 'show_arrow': True},
        {'label': 'UNIQUE ITEMS', 'value': str(tot_items),
         'delta': 'Distinct products ordered', 'delta_tone': 'neutral', 'show_arrow': True},
        {'label': 'SR ORDERS COUNT', 'value': str(tot_orders),
         'delta': 'Warehouse vouchers processed', 'delta_tone': 'pos', 'show_arrow': True},
    ])

    # Tabs
    tab_overview, tab_gs, tab_cs, tab_ps, tab_pst, tab_all, tab_exp = st.tabs([
        "📊 Category Overview",
        "🛎️ Guest Supplies",
        "🧹 Cleaning Supplies",
        "🧻 Paper Supplies",
        "📑 Print & Stationery",
        "📋 All Stock Requests Log",
        "📥 Excel Export"
    ])

    cat_order = ['Guest Supplies', 'Cleaning Supplies', 'Paper Supplies', 'Print & Stationery']

    # Tab 1: Category Overview
    with tab_overview:
        col_ch1, col_ch2 = st.columns([1, 1])

        cat_sum = sr_parser.get_category_summary(sr_df)

        with col_ch1:
            st.markdown("#### Spending Share by Category")
            fig_pie = ui_components.themed_donut(cat_sum, names='Category', values='Total_Amount')
            st.plotly_chart(fig_pie, width="stretch")

        with col_ch2:
            st.markdown("#### Top 10 Cost Drivers (All SR Items)")
            top_10 = sr_df.groupby(['Item_Name', 'Category'], as_index=False)['Total'].sum().sort_values(by='Total', ascending=True).tail(10)
            fig_bar = ui_components.themed_bar(
                top_10,
                x='Total',
                y='Item_Name',
                color='Category',
                orientation='h',
                value_axis_title="Total Spend (IDR)",
                value_tickformat=',',
                showlegend=False,
            )
            st.plotly_chart(fig_bar, width="stretch")

        st.markdown("#### Category Breakdown Summary")
        disp_cat = cat_sum.copy()
        disp_cat['Total_Qty'] = disp_cat['Total_Qty'].apply(lambda x: f"{x:,.0f}")
        disp_cat['Total_Amount'] = disp_cat['Total_Amount'].apply(format_idr)
        disp_cat['Pct_Of_Total'] = disp_cat['Pct_Of_Total'].apply(lambda x: f"{x:.1f}%")
        st.dataframe(disp_cat, width="stretch", hide_index=True)

    # Helper function for rendering a category tab
    def render_category_view(category_name: str, emoji: str):
        c_items = sr_parser.get_item_summary(sr_df, category_name)
        c_trxs = sr_df[sr_df['Category'] == category_name].sort_values(by=['Date', 'Item_Name'])

        if c_items.empty:
            st.info(f"No {category_name} items recorded in the uploaded SR files.")
            return

        cat_amt = c_items['Total_Amount'].sum()
        cat_q = c_items['Total_Qty'].sum()
        pct_all = (cat_amt / tot_amt * 100.0) if tot_amt > 0 else 0.0

        st.markdown(f"### {emoji} {category_name}")
        st.markdown(f"**Total Spend:** {format_idr(cat_amt)} ({pct_all:.1f}% of SR spend) • **Total Qty:** {cat_q:,.0f} • **Unique Items:** {len(c_items)}")

        st.markdown("#### Item Totals Summary (Repeated Orders Aggregated)")
        disp_items = c_items.copy()
        disp_items['Total_Qty'] = disp_items['Total_Qty'].apply(lambda x: f"{x:,.0f}")
        disp_items['Avg_Cost'] = disp_items['Avg_Cost'].apply(format_idr)
        disp_items['Total_Amount'] = disp_items['Total_Amount'].apply(format_idr)
        st.dataframe(disp_items, width="stretch", hide_index=True)

        with st.expander(f"📄 View Chronological Order & Issuing Log ({len(c_trxs)} records)", expanded=False):
            disp_trxs = c_trxs[['Date', 'SR_Number', 'Item_Code', 'Item_Name', 'Qty', 'Unit', 'Cost', 'Total', 'Requested_By']].copy()
            disp_trxs['Qty'] = disp_trxs['Qty'].apply(lambda x: f"{x:,.0f}")
            disp_trxs['Cost'] = disp_trxs['Cost'].apply(format_idr)
            disp_trxs['Total'] = disp_trxs['Total'].apply(format_idr)
            st.dataframe(disp_trxs, width="stretch", hide_index=True)

    with tab_gs:
        render_category_view('Guest Supplies', '🛎️')

    with tab_cs:
        render_category_view('Cleaning Supplies', '🧹')

    with tab_ps:
        render_category_view('Paper Supplies', '🧻')

    with tab_pst:
        render_category_view('Print & Stationery', '📑')

    with tab_all:
        st.markdown("### 📋 Complete Stock Request Log")
        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            cat_filter = st.selectbox("Filter by Category", ["All Categories"] + cat_order)
        with col_f2:
            search_query = st.text_input("Search item, SR number, or requester", "")

        filtered = sr_df.copy()
        if cat_filter != "All Categories":
            filtered = filtered[filtered['Category'] == cat_filter]
        if search_query:
            q = re.escape(search_query.lower())
            filtered = filtered[
                filtered['Item_Name'].str.lower().str.contains(q) |
                filtered['SR_Number'].str.lower().str.contains(q) |
                filtered['Requested_By'].str.lower().str.contains(q)
            ]

        disp_all = filtered[['Date', 'SR_Number', 'Category', 'Item_Code', 'Item_Name', 'Qty', 'Unit', 'Cost', 'Total', 'Requested_By']].copy()
        disp_all['Cost'] = disp_all['Cost'].apply(format_idr)
        disp_all['Total'] = disp_all['Total'].apply(format_idr)
        st.dataframe(disp_all, width="stretch", hide_index=True)

        csv_bytes = filtered.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Filtered Log (CSV)", csv_bytes, "Stock_Requests_Filtered.csv", "text/csv")

    with tab_exp:
        st.markdown("### 📥 Excel Export - Separated Stock Request Report")
        st.markdown("""
        The generated Excel workbook contains multiple pre-styled tabs matching the hotel standard format:
        - 📊 **SR Summary**: KPI metrics, category spend breakdown, % shares, and top 10 cost drivers.
        - 🛎️ **Guest Supplies**: Item summary totals + chronological issuing history.
        - 🧹 **Cleaning Supplies**: Chemical and trash bag item summary + order log.
        - 🧻 **Paper Supplies**: Facial tissue, hand towel, and toilet tissue totals + order log.
        - 📑 **Print & Stationery**: Form and office supply totals + order log.
        - 📋 **All SR Items**: Complete line item audit table.
        """)

        st.download_button(
            label=f"📥 Download Full Separated SR Excel ({sr_report_title})",
            data=excel_sr_bytes,
            file_name=f"Separated_SR_{sr_report_title.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )


# Base directories for standard data
if "selected_data_dir" not in st.session_state:
    st.session_state.selected_data_dir = get_default_data_dir()

BASE_REVIEW_DIR = st.session_state.selected_data_dir
st.session_state.selected_data_dir = os.path.normpath(BASE_REVIEW_DIR)

MONTH_FOLDERS = discover_month_folders(BASE_REVIEW_DIR)
if not MONTH_FOLDERS:
    MONTH_FOLDERS = {
        "August 2026 (Agustus)": os.path.join(BASE_REVIEW_DIR, "8.AGUSTUS"),
        "July 2026 (Juli)": os.path.join(BASE_REVIEW_DIR, "7.JULY"),
    }
has_local_folder = os.path.exists(BASE_REVIEW_DIR)

st.sidebar.markdown("### 📁 Data Folder")
user_data_dir = st.sidebar.text_input(
    "Folder with month data",
    value=st.session_state.selected_data_dir,
    help="Set the parent folder containing month folders such as 8.AGUSTUS/ and 7.JULY/."
)
if user_data_dir:
    st.session_state.selected_data_dir = user_data_dir
    BASE_REVIEW_DIR = user_data_dir
    MONTH_FOLDERS = discover_month_folders(BASE_REVIEW_DIR)
    if not MONTH_FOLDERS:
        MONTH_FOLDERS = {
            "August 2026 (Agustus)": os.path.join(BASE_REVIEW_DIR, "8.AGUSTUS"),
            "July 2026 (Juli)": os.path.join(BASE_REVIEW_DIR, "7.JULY"),
        }
    has_local_folder = os.path.exists(BASE_REVIEW_DIR)

if not has_local_folder:
    st.sidebar.warning(f"No local data folder found at: {BASE_REVIEW_DIR}")
    st.sidebar.caption("Set a valid folder with 8.AGUSTUS/ and 7.JULY/, or use demo data.")

# Sidebar navigation & data selection
app_mode = st.sidebar.radio(
    "Application Mode",
    ["💰 Monthly Expense Tracker", "📦 Stock Request (SR) Separator"],
    index=0
)
st.sidebar.divider()

if "demo_mode" not in st.session_state:
    st.session_state.demo_mode = False

if st.session_state.demo_mode:
    st.sidebar.success("Demo data is active")
    if st.sidebar.button("Reset to Local Folder Mode", type="secondary", width="stretch", key="reset_demo"):
        st.session_state.demo_mode = False
        st.rerun()
else:
    if st.sidebar.button("Load Demo Data", type="secondary", width="stretch", key="load_demo"):
        st.session_state.demo_mode = True
        st.rerun()

# Load Sample Data button - adds a preset demo dataset for testing
if st.sidebar.button("Load Sample Data", type="secondary", width="stretch", key="load_sample"):
    st.session_state.demo_mode = True
    # Create a simple demo dataset using the existing demo builder
    demo_data = build_demo_reconciled_data()
    st.session_state.uploaded_months = {"Sample August 2026": demo_data}
    st.session_state.upload_counter = 1
    st.toast("Sample data loaded!")
    st.rerun()

# Keyboard shortcuts
st.markdown("""
<script>
document.addEventListener('keydown', function(e) {
    if (e.ctrlKey && e.key === 'r') {
        e.preventDefault();
        // Trigger reset by clicking the reset button if it exists
        const resetBtn = document.querySelector('button[kind="secondary"]');
        if (resetBtn) resetBtn.click();
    }
});
</script>
""", unsafe_allow_html=True)

if app_mode == "📦 Stock Request (SR) Separator":
    render_sr_separator()
    st.stop()

# --- Monthly Expense Tracker Mode Below ---
if st.session_state.demo_mode:
    data_source = "Demo Data"
else:
    data_source = st.sidebar.radio(
        "Select Data Source",
        ["Select Month Folder", "Upload Custom Excel Files", "Upload Multiple Months"],
        index=0 if has_local_folder else 1
    )

reconciled_data = None
selected_month_name = "Demo August 2026"

if data_source == "Demo Data":
    reconciled_data = build_demo_reconciled_data()
    selected_month_name = "Demo August 2026"
    st.sidebar.success("Demo data loaded. This is a sample month for previewing the dashboard.")

elif data_source == "Select Month Folder":
    selected_months = st.sidebar.multiselect(
        "Choose Month(s) to Load",
        options=list(MONTH_FOLDERS.keys()),
        default=list(MONTH_FOLDERS.keys()) if has_local_folder else []
    )
    if selected_months:
        month_data_dict = {}
        validation_messages = []
        for m_choice in selected_months:
            m_dir = MONTH_FOLDERS[m_choice]
            m_label = parser.infer_month_label(m_choice)
            if not os.path.exists(m_dir):
                validation_messages.append(f"Folder not found for {m_label}: {m_dir}")
                continue
            files = parser.find_month_files(m_dir)
            if not files['dtb']:
                validation_messages.append(f"Detail Trial Balance not found for {m_label}; skipped.")
                continue
            with st.spinner(f"Parsing {m_label}..."):
                dtb_path = files['dtb']
                is_path = files['is_mtd']
                cons_path = files['consumption']
                import tempfile
                cached_data = cached_reconcile_month(dtb_path, is_path, cons_path, m_label)
                month_data_dict[m_label] = cached_data
        if validation_messages:
            show_validation_banner(validation_messages)
        if month_data_dict:
            reconciled_data, selected_month_name = combine_month_datasets(month_data_dict)
        else:
            st.sidebar.warning("No valid month data could be loaded from the selected folders.")
    else:
        st.sidebar.warning("Please select at least one month folder.")

elif data_source == "Upload Custom Excel Files":
    if "uploaded_months" not in st.session_state:
        st.session_state.uploaded_months = {}
    if "upload_counter" not in st.session_state:
        st.session_state.upload_counter = 0

    # Display currently loaded months with delete buttons
    if st.session_state.uploaded_months:
        st.sidebar.markdown("### 📁 Loaded Month Datasets")
        for m_key in list(st.session_state.uploaded_months.keys()):
            c1, c2 = st.sidebar.columns([4, 1])
            c1.write(f"• **{m_key}**")
            if c2.button("❌", key=f"del_{m_key}"):
                del st.session_state.uploaded_months[m_key]
                st.rerun()
        if st.sidebar.button("🗑️ Clear All Loaded Months", type="secondary", width="stretch"):
            st.session_state.uploaded_months = {}
            st.rerun()
        st.sidebar.divider()

    st.sidebar.markdown("**Add Month Files:**")
    default_month_label = f"Month {len(st.session_state.uploaded_months) + 1}"
    selected_month_label = st.sidebar.text_input("Month Label (e.g. July 2026, Agustus 2026)", default_month_label)

    dtb_file = st.sidebar.file_uploader("1. Detail Trial Balance (DTB)", type=["xlsx"], key=f"dtb_{st.session_state.upload_counter}")
    is_file = st.sidebar.file_uploader("2. Income Statement Dept (MTD)", type=["xlsx"], key=f"is_{st.session_state.upload_counter}")
    cons_file = st.sidebar.file_uploader("3. Consumption Report", type=["xlsx"], key=f"cons_{st.session_state.upload_counter}")

    if st.sidebar.button("➕ Add This Month to Dashboard", type="primary", width="stretch"):
        validation_messages = []
        if not dtb_file:
            validation_messages.append("Detail Trial Balance (DTB) file is required.")
        if selected_month_label in st.session_state.uploaded_months:
            validation_messages.append(f"Month label '{selected_month_label}' already exists — rename it to add another month.")
        
        # Basic filename validation to prevent path traversal
        for uploaded_file in [dtb_file, is_file, cons_file]:
            if uploaded_file and (".." in uploaded_file.name or "/" in uploaded_file.name or "\\" in uploaded_file.name):
                validation_messages.append(f"Invalid filename detected: {uploaded_file.name}. Filenames cannot contain path separators or '..'.")

        if validation_messages:
            show_validation_banner(validation_messages, level="error")
        else:
            with st.spinner(f"Processing {selected_month_label}..."):
                import tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f_dtb:
                    f_dtb.write(dtb_file.read())
                    dtb_path = f_dtb.name

                is_path = None
                if is_file:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f_is:
                        f_is.write(is_file.read())
                        is_path = f_is.name

                cons_path = None
                if cons_file:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f_cons:
                        f_cons.write(cons_file.read())
                        cons_path = f_cons.name

                dtb_df = parser.parse_detail_trial_balance(dtb_path)
                is_df = parser.parse_income_statement(is_path) if is_path else pd.DataFrame()
                cons_df = parser.parse_consumption_report(cons_path) if cons_path else pd.DataFrame()
                parsed_m_data = matcher.reconcile_monthly_expenses(dtb_df, is_df, cons_df)

                st.session_state.uploaded_months[selected_month_label] = parsed_m_data
                st.session_state.upload_counter += 1

                try:
                    os.unlink(dtb_path)
                    if is_path: os.unlink(is_path)
                    if cons_path: os.unlink(cons_path)
                except Exception:
                    pass

                st.toast(f"Added {selected_month_label} successfully!")
                st.rerun()

    # Automatically combine all accumulated uploaded months
    if st.session_state.uploaded_months:
        reconciled_data, selected_month_name = combine_month_datasets(st.session_state.uploaded_months)

elif data_source == "Upload Multiple Months":
    zip_file = st.sidebar.file_uploader("Upload ZIP of month folders", type=["zip"])
    if zip_file:
        import zipfile, tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(zip_file) as z:
                z.extractall(tmpdir)
            month_dirs = [os.path.join(tmpdir, d) for d in os.listdir(tmpdir) if os.path.isdir(os.path.join(tmpdir, d))]
            month_data = {}
            for month_dir in month_dirs:
                raw_label = os.path.basename(month_dir)
                month_label = parser.infer_month_label(raw_label, default=raw_label)
                files = parser.find_month_files(month_dir)
                if not files['dtb']:
                    st.warning(f"DTB not found in {month_label}, skipping.")
                    continue
                with st.spinner(f"Processing {month_label}..."):
                    dtb_df = parser.parse_detail_trial_balance(files['dtb'])
                    is_df = parser.parse_income_statement(files['is_mtd']) if files['is_mtd'] else pd.DataFrame()
                    cons_df = parser.parse_consumption_report(files['consumption']) if files['consumption'] else pd.DataFrame()
                    month_data[month_label] = matcher.reconcile_monthly_expenses(dtb_df, is_df, cons_df)
            if month_data:
                reconciled_data, selected_month_name = combine_month_datasets(month_data)

# Label used for MoM tab heading (falls back sensibly for multi-month selections)
mom_month_name = selected_month_name


if reconciled_data:
    metrics = reconciled_data['metrics']
    cat_df = reconciled_data['category_summary']
    trx_df = reconciled_data['transactions']
    top_items = reconciled_data['top_cost_drivers']
    item_df = reconciled_data.get('item_summary', pd.DataFrame())

    # Pre-generate Excel export buffer for sidebar one-click download
    excel_bytes = exporter.create_separated_excel(reconciled_data, selected_month_name)
    st.sidebar.divider()
    st.sidebar.markdown("### 📥 Quick Export")
    st.sidebar.download_button(
        label="Download Separated Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"Separated_Expenses_{selected_month_name.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        type="primary"
    )

    # Main Page Header
    st.markdown('<div class="main-header">Hotel Expense Separator & Spending Tracker</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Hotel Santika Depok • Room Division & Housekeeping • Period: <b>{selected_month_name}</b></div>', unsafe_allow_html=True)

    # Top KPI Metrics row
    budget_pct = (metrics['total_spent'] / metrics['total_budget'] * 100) if metrics['total_budget'] > 0 else 0
    ui_components.render_kpi_row([
        {'label': 'TOTAL SPENT', 'value': format_idr(metrics['total_spent']), 'delta': None, 'show_arrow': False},
        {'label': 'TOTAL BUDGET', 'value': format_idr(metrics['total_budget']), 'delta': None, 'show_arrow': False},
        {'label': 'NET VARIANCE', 'value': format_idr(metrics['variance_idr']),
         'delta': f"{metrics['variance_pct']:+.1f}% vs budget",
         'delta_tone': 'pos' if metrics['variance_idr'] <= 0 else 'neg', 'show_arrow': False},
        {'label': 'RECONCILED TRANSACTIONS', 'value': f"{metrics['transaction_count']:,}",
         'delta': 'ledger entries classified', 'delta_tone': 'neutral', 'show_arrow': False},
        {'label': 'BUDGET UTILIZATION', 'value': f"{budget_pct:.1f}%",
         'delta': 'share of budget consumed', 'delta_tone': 'neutral', 'show_arrow': False},
    ])

    # Tabs
    tab_dash, tab_items, tab_drill, tab_bgt, tab_mom, tab_export = st.tabs([
        "📊 Executive Dashboard",
        "📦 Item Totals & Quantities",
        "🔍 Transaction Drilldown",
        "⚖️ Budget vs Actual Variance",
        "📈 Month-over-Month Comparison",
        "📂 Separated Excel & Export"
    ])

    # ==========================================
    # TAB 1: EXECUTIVE DASHBOARD
    # ==========================================
    with tab_dash:
        col_left, col_right = st.columns([1, 1])

        with col_left:
            st.subheader("Spend Distribution by Category")
            fig_pie = ui_components.themed_donut(
                cat_df, values="Actual", names="Category", height=360, showlegend=False
            )
            st.plotly_chart(fig_pie, width="stretch")

        with col_right:
            st.subheader("Budget vs. Actual (Top 8 Categories)")
            top_cats = cat_df.head(8).sort_values(by="Actual", ascending=True)
            fig_bar = ui_components.themed_comparison_bar(
                top_cats['Category'].tolist(),
                [
                    {'name': 'Budget', 'values': top_cats['Budget'].tolist()},
                    {'name': 'Actual', 'values': top_cats['Actual'].tolist()},
                ],
                orientation='h',
                height=360,
                legend_horizontal=True,
            )
            st.plotly_chart(fig_bar, width="stretch")

        st.divider()

        st.subheader("Top 10 Largest Cost Drivers Across Hotel")
        top_10 = top_items.head(10).sort_values(by="Total_Amount", ascending=True)
        fig_cost = ui_components.themed_bar(
            top_10,
            x="Total_Amount",
            y="Item_Name",
            color="Category",
            orientation="h",
            labels={"Total_Amount": "Amount (IDR)", "Item_Name": "Expense Item / Vendor"},
            height=400,
            value_tickformat=',',
        )
        st.plotly_chart(fig_cost, width="stretch")

    # ==========================================
    # TAB: ITEM CONSUMPTION & REPEATED ORDERS
    # ==========================================
    with tab_items:
        st.subheader("📦 Item Consumption & Repeated Orders Summary")
        st.markdown(
            "Track monthly cumulative quantities, average unit prices, and total spend per item "
            "across recurring orders (e.g. **Cleo 330ml mineral water**, **Toilet Tissue**, **Slippers**, **Soap**, **Garbage Bags**)."
        )

        # Filters
        c_filter_col1, c_filter_col2 = st.columns([2, 3])
        
        available_cats = sorted(item_df['Category'].unique().tolist()) if not item_df.empty else []
        cat_choices = ["All Supplies Categories", "All Categories"] + available_cats
        seen = set()
        cat_choices = [x for x in cat_choices if not (x in seen or seen.add(x))]

        with c_filter_col1:
            sel_item_cat = st.selectbox("Select Expense Category", cat_choices)
        with c_filter_col2:
            item_search = st.text_input("🔍 Search Item Name (e.g. Cleo, Tissue, Slipper, Soap)", "")

        filtered_items = item_df.copy() if not item_df.empty else pd.DataFrame()
        if not filtered_items.empty:
            if sel_item_cat == "All Supplies Categories":
                filtered_items = filtered_items[filtered_items['Category'].isin([
                    'Guest Supplies', 'Paper Suplies', 'Cleaning Supplies', 'Printing & Stationery,Photo Copy, Postage & Stamp'
                ])]
            elif sel_item_cat != "All Categories":
                filtered_items = filtered_items[filtered_items['Category'] == sel_item_cat]

            if item_search:
                filtered_items = filtered_items[filtered_items['Item_Name'].str.lower().str.contains(re.escape(item_search.lower()), na=False)]

        if not filtered_items.empty:
            # Summary KPIs for selected items
            total_items_spend = filtered_items['Total_Amount'].sum()
            total_items_qty = filtered_items['Total_Qty'].sum()
            total_orders_count = filtered_items['Order_Count'].sum()
            unique_items_cnt = len(filtered_items)

            ui_components.render_kpi_row([
                {'label': 'TOTAL SPEND', 'value': format_idr(total_items_spend), 'delta': None, 'show_arrow': False},
                {'label': 'TOTAL UNITS / QUANTITIES', 'value': f"{total_items_qty:,.0f} units", 'delta': None, 'show_arrow': False},
                {'label': 'TOTAL REPEATED ORDERS', 'value': str(total_orders_count), 'delta': 'deliveries', 'delta_tone': 'neutral', 'show_arrow': True},
                {'label': 'UNIQUE ITEMS TRACKED', 'value': str(unique_items_cnt), 'delta': 'items', 'delta_tone': 'neutral', 'show_arrow': True},
            ])

            # Item Summary Table
            st.markdown("#### Item Totals Table (Quantities & Total Cost)")
            display_item_df = filtered_items[[
                'Category', 'Item_Name', 'Total_Qty', 'Unit',
                'Avg_Unit_Price', 'Total_Amount', 'Pct_Of_Category', 'Order_Count', 'First_Date', 'Last_Date'
            ]].copy()

            display_item_df['Total_Quantity'] = display_item_df.apply(
                lambda r: f"{r['Total_Qty']:,.0f} {r['Unit']}".strip(), axis=1
            )
            display_item_df['Avg_Unit_Price'] = display_item_df['Avg_Unit_Price'].apply(format_idr)
            display_item_df['Total_Amount'] = display_item_df['Total_Amount'].apply(format_idr)
            display_item_df['Pct_Of_Category'] = display_item_df['Pct_Of_Category'].apply(format_pct)
            display_item_df['Delivery_Period'] = display_item_df.apply(
                lambda r: f"{r['First_Date']} ~ {r['Last_Date']}" if r['First_Date'] != r['Last_Date'] else str(r['First_Date']),
                axis=1
            )
            
            table_view = display_item_df[[
                'Category', 'Item_Name', 'Total_Quantity', 'Avg_Unit_Price',
                'Total_Amount', 'Pct_Of_Category', 'Order_Count', 'Delivery_Period'
            ]]

            st.dataframe(
                table_view,
                width="stretch",
                height=380,
                column_config={
                    "Category": st.column_config.TextColumn("Category", width="medium"),
                    "Item_Name": st.column_config.TextColumn("Item / Product Name", width="large"),
                    "Total_Quantity": st.column_config.TextColumn("Total Quantity", width="medium"),
                    "Avg_Unit_Price": st.column_config.TextColumn("Avg Unit Price", width="small"),
                    "Total_Amount": st.column_config.TextColumn("Total Spend (IDR)", width="medium"),
                    "Pct_Of_Category": st.column_config.TextColumn("% of Category", width="small"),
                    "Order_Count": st.column_config.NumberColumn("Orders", width="small"),
                    "Delivery_Period": st.column_config.TextColumn("Delivery Period", width="medium"),
                }
            )

            st.divider()

            # Item Inspector: Deep dive into individual deliveries
            st.markdown("#### 🔬 Item Inspector: Delivery & Order History")
            item_options = filtered_items['Item_Name'].unique().tolist()
            if item_options:
                default_idx = 0
                for idx, name in enumerate(item_options):
                    if "cleo" in name.lower():
                        default_idx = idx
                        break

                selected_item_name = st.selectbox(
                    "Select an item to view its complete delivery timeline and voucher log:",
                    item_options,
                    index=default_idx
                )

                item_trxs = trx_df[trx_df['Item_Name'] == selected_item_name].sort_values(by='Date')
                
                if not item_trxs.empty:
                    item_tot_qty = item_trxs['Qty'].sum()
                    item_unit = item_trxs['Unit'].iloc[0] or 'units'
                    item_tot_cost = item_trxs['Amount'].sum()
                    item_avg_cost = item_trxs['Unit_Price'].mean()

                    ui_components.render_kpi_row([
                        {'label': f'TOTAL {selected_item_name.upper()}', 'value': f"{item_tot_qty:,.0f} {item_unit}", 'delta': None, 'show_arrow': False},
                        {'label': 'TOTAL MONTH SPEND', 'value': format_idr(item_tot_cost), 'delta': None, 'show_arrow': False},
                        {'label': 'AVG UNIT PRICE', 'value': format_idr(item_avg_cost), 'delta': None, 'show_arrow': False},
                        {'label': 'TOTAL DELIVERIES / VOUCHERS', 'value': str(len(item_trxs)), 'delta': 'times', 'delta_tone': 'neutral', 'show_arrow': True},
                    ])

                    # Timeline Chart of Deliveries
                    fig_timeline = ui_components.themed_bar(
                        item_trxs,
                        x='Date',
                        y='Qty',
                        hover_data=['Amount', 'Voucher_Ref', 'Unit_Price'],
                        title=f"Delivery Timeline for {selected_item_name} ({item_unit})",
                        labels={'Qty': f"Quantity Delivered ({item_unit})", 'Date': "Order / Issuing Date"},
                        height=280,
                        margin_top=35,
                    )
                    st.plotly_chart(fig_timeline, width="stretch")

                    # Detailed Voucher Table
                    st.markdown(f"**Individual Order Vouchers for {selected_item_name}:**")
                    voucher_view = item_trxs[['Date', 'Voucher_Ref', 'Qty', 'Unit', 'Unit_Price', 'Amount', 'JRNL', 'Raw_Description']].copy()
                    voucher_view['Unit_Price'] = voucher_view['Unit_Price'].apply(format_idr)
                    voucher_view['Amount'] = voucher_view['Amount'].apply(format_idr)
                    st.dataframe(voucher_view, width="stretch")

            # Download Item Summary CSV
            item_csv = filtered_items.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"📥 Download Item Totals Summary (CSV)",
                data=item_csv,
                file_name=f"Item_Totals_{sel_item_cat.replace(' ', '_')}.csv",
                mime="text/csv"
            )
        else:
            st.info("No items match the selected criteria.")

    # ==========================================
    # TAB 3: TRANSACTION DRILLDOWN
    # ==========================================
    with tab_drill:
        st.subheader("Detailed Spending Explorer")
        st.markdown("Inspect every individual transaction, vendor voucher, item quantity, and line amount.")

        # Filters
        fcol1, fcol2, fcol3 = st.columns([2, 2, 2])
        with fcol1:
            all_cats = ["All Categories"] + sorted(trx_df['Category'].unique().tolist())
            selected_cat = st.selectbox("Filter by Category", all_cats)
        with fcol2:
            search_query = st.text_input("🔍 Search Item / Vendor / Voucher", "")
        with fcol3:
            sort_by = st.selectbox("Sort By", ["Amount (High to Low)", "Amount (Low to High)", "Date (Latest)", "Item Name"])

        # Filter logic
        filtered_df = trx_df.copy()
        if selected_cat != "All Categories":
            filtered_df = filtered_df[filtered_df['Category'] == selected_cat]

        if search_query:
            q = re.escape(search_query.lower())
            filtered_df = filtered_df[
                filtered_df['Item_Name'].str.lower().str.contains(q, na=False) |
                filtered_df['Partner_Vendor'].str.lower().str.contains(q, na=False) |
                filtered_df['Voucher_Ref'].str.lower().str.contains(q, na=False) |
                filtered_df['Raw_Description'].str.lower().str.contains(q, na=False)
            ]

        # Sorting
        if sort_by == "Amount (High to Low)":
            filtered_df = filtered_df.sort_values(by="Amount", ascending=False)
        elif sort_by == "Amount (Low to High)":
            filtered_df = filtered_df.sort_values(by="Amount", ascending=True)
        elif sort_by == "Date (Latest)":
            filtered_df = filtered_df.sort_values(by="Date", ascending=False)
        elif sort_by == "Item Name":
            filtered_df = filtered_df.sort_values(by="Item_Name", ascending=True)

        st.caption(f"Showing {len(filtered_df)} transactions totaling **{format_idr(filtered_df['Amount'].sum())}**")

        # Display Data Table
        display_df = filtered_df[[
            'Date', 'Category', 'Item_Name', 'Partner_Vendor',
            'Qty', 'Unit', 'Unit_Price', 'Amount', 'Voucher_Ref', 'JRNL'
        ]].copy()
        display_df['Unit_Price'] = display_df['Unit_Price'].apply(lambda v: f"Rp {v:,.0f}")
        display_df['Amount'] = display_df['Amount'].apply(lambda v: f"Rp {v:,.0f}")

        st.dataframe(
            display_df,
            width="stretch",
            height=480,
            column_config={
                "Date": st.column_config.TextColumn("Date", width="small"),
                "Category": st.column_config.TextColumn("Category", width="medium"),
                "Item_Name": st.column_config.TextColumn("Item / Service Description", width="large"),
                "Partner_Vendor": st.column_config.TextColumn("Vendor / Partner", width="medium"),
                "Qty": st.column_config.NumberColumn("Qty", format="%.0f"),
                "Unit": st.column_config.TextColumn("Unit", width="small"),
                "Unit_Price": st.column_config.TextColumn("Unit Price", width="small"),
                "Amount": st.column_config.TextColumn("Total Amount", width="medium"),
                "Voucher_Ref": st.column_config.TextColumn("Ref / Voucher", width="medium"),
            }
        )

        # Download current slice
        csv_slice = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Filtered Results (CSV)",
            data=csv_slice,
            file_name=f"Filtered_Expenses_{selected_cat.replace(' ', '_')}.csv",
            mime="text/csv"
        )

    # ==========================================
    # TAB 3: BUDGET VS ACTUAL VARIANCE
    # ==========================================
    with tab_bgt:
        st.subheader("Budget vs. Actual Variance Analysis")
        st.markdown("Identify budget overruns and operational savings at a glance.")

        # Variance alerts for over-budget categories
        over_bgt = cat_df[cat_df['Status'] == 'Over Budget']
        if not over_bgt.empty:
            st.warning(f"⚠️ **Attention:** {len(over_bgt)} categories exceeded their budget allocation this month!")
            cols = st.columns(min(len(over_bgt), 3))
            for i, (_, row) in enumerate(over_bgt.iterrows()):
                with cols[i % len(cols)]:
                    st.error(f"**{row['Category']}**\n\nActual: {format_idr(row['Actual'])}\n\nBudget: {format_idr(row['Budget'])}\n\n**Over by {format_idr(row['Variance_IDR'])} (+{row['Variance_Pct']:.1f}%)**")

        st.markdown("<br>", unsafe_allow_html=True)

        # Summary Variance Table
        var_display = cat_df[[
            'Category', 'Group', 'Budget', 'Actual', 'Variance_IDR', 'Variance_Pct', 'Status', 'Top_Item', 'Transaction_Count'
        ]].copy()
        
        var_display['Budget'] = var_display['Budget'].apply(format_idr)
        var_display['Actual'] = var_display['Actual'].apply(format_idr)
        var_display['Variance_IDR'] = var_display['Variance_IDR'].apply(format_idr)
        var_display['Variance_Pct'] = var_display['Variance_Pct'].apply(format_pct)

        st.dataframe(
            var_display,
            width="stretch",
            height=450,
            column_config={
                "Category": st.column_config.TextColumn("Expense Category", width="large"),
                "Group": st.column_config.TextColumn("Group", width="small"),
                "Budget": st.column_config.TextColumn("Budget (IDR)"),
                "Actual": st.column_config.TextColumn("Actual (IDR)"),
                "Variance_IDR": st.column_config.TextColumn("Variance (IDR)"),
                "Variance_Pct": st.column_config.TextColumn("Variance (%)"),
                "Status": st.column_config.TextColumn("Status", width="small"),
                "Top_Item": st.column_config.TextColumn("Top Cost Item", width="large"),
                "Transaction_Count": st.column_config.NumberColumn("Txns", width="small"),
            }
        )

    # ==========================================
    # TAB 4: MONTH-OVER-MONTH COMPARISON
    # ==========================================
    with tab_mom:
        st.subheader("Month-over-Month Spending Comparison")
        mom_folder_pairs = list(MONTH_FOLDERS.items())
        if len(mom_folder_pairs) >= 2:
            mom_current_label, mom_current_dir = mom_folder_pairs[0]
            mom_previous_label, mom_previous_dir = mom_folder_pairs[1]
            mom_current_name = parser.infer_month_label(mom_current_label)
            mom_previous_name = parser.infer_month_label(mom_previous_label)
            st.markdown(
                f"Track expense shifts between **{mom_current_name}** "
                f"and **{mom_previous_name}**."
            )
        else:
            mom_current_label = mom_previous_label = None
            mom_current_name = mom_previous_name = None
            mom_current_dir = mom_previous_dir = None
            st.info("Not enough month folders configured for MoM comparison.")

        if mom_current_dir and mom_previous_dir:
            mom_current_files = parser.find_month_files(mom_current_dir)
            mom_previous_files = parser.find_month_files(mom_previous_dir)

            if not mom_previous_files['dtb']:
                st.info("Previous month data not found for MoM comparison.")
            else:
                # If the currently selected month isn't the MoM current month,
                # re-parse it directly from its folder for a like-for-like comparison.
                if (selected_month_name == mom_current_name
                        and mom_current_files['dtb']):
                    cur_dtb = parser.parse_detail_trial_balance(mom_current_files['dtb'])
                    cur_is = parser.parse_income_statement(mom_current_files['is_mtd']) if mom_current_files['is_mtd'] else pd.DataFrame()
                    cur_cons = parser.parse_consumption_report(mom_current_files['consumption']) if mom_current_files['consumption'] else pd.DataFrame()
                    cur_rec = matcher.reconcile_monthly_expenses(cur_dtb, cur_is, cur_cons)
                else:
                    cur_rec = reconciled_data

                with st.spinner("Calculating MoM trends..."):
                    j_dtb = parser.parse_detail_trial_balance(mom_previous_files['dtb'])
                    j_is = parser.parse_income_statement(mom_previous_files['is_mtd']) if mom_previous_files['is_mtd'] else pd.DataFrame()
                    j_cons = parser.parse_consumption_report(mom_previous_files['consumption']) if mom_previous_files['consumption'] else pd.DataFrame()
                    j_rec = matcher.reconcile_monthly_expenses(j_dtb, j_is, j_cons)

                if cur_rec['metrics']['transaction_count'] == 0 or j_rec['metrics']['transaction_count'] == 0:
                    st.info("Previous month data not found for MoM comparison.")
                else:
                    a_cat = cur_rec['category_summary'][['Category', 'Actual']].rename(columns={'Actual': 'August_Actual'})
                    j_cat = j_rec['category_summary'][['Category', 'Actual']].rename(columns={'Actual': 'July_Actual'})

                    mom_df = pd.merge(a_cat, j_cat, on="Category", how="outer").fillna(0.0)
                    mom_df['MoM_Diff'] = mom_df['August_Actual'] - mom_df['July_Actual']
                    mom_df['MoM_Pct'] = (mom_df['MoM_Diff'] / mom_df['July_Actual'] * 100.0).replace([np.inf, -np.inf], 0.0).fillna(0.0)
                    mom_df = mom_df.sort_values(by="August_Actual", ascending=False)

                    # MoM KPI overview
                    july_total = j_rec['metrics']['total_spent']
                    net_mom = cur_rec['metrics']['total_spent'] - july_total
                    net_mom_pct = (net_mom / july_total * 100.0) if july_total > 0 else 0
                    ui_components.render_kpi_row([
                        {'label': (mom_current_name + ' Total').upper(), 'value': format_idr(cur_rec['metrics']['total_spent']), 'delta': None, 'show_arrow': False},
                        {'label': (mom_previous_name + ' Total').upper(), 'value': format_idr(july_total), 'delta': None, 'show_arrow': False},
                        {'label': 'MOM SPENDING SHIFT', 'value': format_idr(net_mom),
                         'delta': f"{net_mom_pct:+.1f}% vs previous month",
                         'delta_tone': 'neg' if net_mom > 0 else 'pos', 'show_arrow': False},
                    ])

                    # MoM Comparison Bar Chart
                    st.markdown("#### Top Categories Comparison")
                    mom_top = mom_df.head(8)
                    fig_mom = ui_components.themed_comparison_bar(
                        mom_top['Category'].tolist(),
                        [
                            {'name': mom_previous_name, 'values': mom_top['July_Actual'].tolist()},
                            {'name': mom_current_name, 'values': mom_top['August_Actual'].tolist()},
                        ],
                        orientation='v',
                        height=380,
                    )
                    st.plotly_chart(fig_mom, width="stretch")

                    # MoM Table
                    mom_display = mom_df.copy()
                    mom_display['August_Actual'] = mom_display['August_Actual'].apply(format_idr)
                    mom_display['July_Actual'] = mom_display['July_Actual'].apply(format_idr)
                    mom_display['MoM_Diff'] = mom_display['MoM_Diff'].apply(format_idr)
                    mom_display['MoM_Pct'] = mom_display['MoM_Pct'].apply(format_pct)

                    st.dataframe(
                        mom_display,
                        width="stretch",
                        height=350,
                        column_config={
                            "Category": st.column_config.TextColumn("Expense Category", width="large"),
                            "July_Actual": st.column_config.TextColumn(mom_previous_name),
                            "August_Actual": st.column_config.TextColumn(mom_current_name),
                            "MoM_Diff": st.column_config.TextColumn("MoM Change (IDR)"),
                            "MoM_Pct": st.column_config.TextColumn("MoM Change (%)"),
                        }
                    )
    # ==========================================
    # TAB 5: SEPARATED EXCEL & EXPORT
    # ==========================================
    with tab_export:
        st.subheader("Separated Excel Workbooks & Downloads")
        st.markdown("""
        The application automatically divides and structures the complex financial data into dedicated sheets:
        - 📑 **Executive Summary**: High-level KPIs, budget vs actual variance, and over/under budget badges.
        - 📑 **All Transactions**: Master filterable transaction journal with item names, vendors, and vouchers.
        - 📑 **Guest Supplies**: Itemized issuing for drinking water (Cleo), amenities, slippers, toothbrush, soap, etc.
        - 📑 **Outsourcing & Laundry**: Vendor breakdowns for Cleaning Service (PT PJT), Bonvivo, and Drop N Go.
        - 📑 **Cleaning Supplies**: Cleaning chemicals, air freshener contracts, hand gloves, garbage bags.
        - 📑 **Paper Supplies**: Toilet rolls, facial tissue, laundry packaging bags.
        - 📑 **Payroll & SC**: Basic salary, daily worker compensation, and service charge distribution.
        - 📑 **Media & Utilities**: Internet connectivity (Maxindo), Indovision Cable TV, Telkom.
        """)

        st.divider()

        st.download_button(
            label=f"📥 Download Full Separated Excel ({selected_month_name})",
            data=excel_bytes,
            file_name=f"Separated_Expenses_{selected_month_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Individual Category Quick Exports (CSV)")
        c1, c2, c3 = st.columns(3)
        with c1:
            gs_csv = trx_df[trx_df['Category'] == 'Guest Supplies'].to_csv(index=False).encode('utf-8')
            st.download_button("Download Guest Supplies (CSV)", gs_csv, "Guest_Supplies.csv", "text/csv")
        with c2:
            out_csv = trx_df[trx_df['Category'].isin(['Outsourcing Utilities', 'Laundry & Dry Cleaning'])].to_csv(index=False).encode('utf-8')
            st.download_button("Download Outsourcing & Laundry (CSV)", out_csv, "Outsourcing_Laundry.csv", "text/csv")
        with c3:
            clean_csv = trx_df[trx_df['Category'].isin(['Cleaning Supplies', 'Paper Suplies'])].to_csv(index=False).encode('utf-8')
            st.download_button("Download Cleaning & Paper (CSV)", clean_csv, "Cleaning_Paper.csv", "text/csv")

else:
    st.info("Please select or upload Excel files from the sidebar to begin.")
