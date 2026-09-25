"""Monthly Expense Tracker view."""
import os
import re
import tempfile
import zipfile

import numpy as np
import pandas as pd
import streamlit as st

import parser
import matcher
import exporter
import ui_components
from ui.formatting import format_idr, format_pct, show_validation_banner
from data.months import (
    get_default_data_dir,
    discover_month_folders,
    cached_reconcile_month,
    build_demo_reconciled_data,
    combine_month_datasets,
)


def _get_month_folders():
    """Return discovered month folders and a flag for local folder existence."""
    base_dir = st.session_state.get("selected_data_dir", get_default_data_dir())
    has_local = os.path.exists(base_dir)
    folders = discover_month_folders(base_dir)
    if not folders:
        folders = {
            "August 2026 (Agustus)": os.path.join(base_dir, "8.AGUSTUS"),
            "July 2026 (Juli)": os.path.join(base_dir, "7.JULY"),
        }
    return base_dir, has_local, folders


def _sidebar_data_dir(base_dir: str, has_local: bool, folders: dict) -> tuple:
    """Render data folder input and update discovered folders."""
    st.sidebar.markdown("### 📁 Data Folder")
    user_data_dir = st.sidebar.text_input(
        "Folder with month data",
        value=base_dir,
        help="Set the parent folder containing month folders such as 8.AGUSTUS/ and 7.JULY/."
    )
    if user_data_dir:
        st.session_state.selected_data_dir = user_data_dir
        base_dir = user_data_dir
        folders = discover_month_folders(base_dir)
        if not folders:
            folders = {
                "August 2026 (Agustus)": os.path.join(base_dir, "8.AGUSTUS"),
                "July 2026 (Juli)": os.path.join(base_dir, "7.JULY"),
            }
        has_local = os.path.exists(base_dir)

    if not has_local:
        st.sidebar.warning(f"No local data folder found at: {base_dir}")
        st.sidebar.caption("Set a valid folder with 8.AGUSTUS/ and 7.JULY/, or use demo data.")
    return base_dir, has_local, folders


def _load_demo_data() -> tuple:
    """Load demo reconciled data."""
    reconciled_data = build_demo_reconciled_data()
    selected_month_name = "Demo August 2026"
    st.sidebar.success("Demo data loaded. This is a sample month for previewing the dashboard.")
    return reconciled_data, selected_month_name


def _load_selected_months(folders: dict) -> tuple:
    """Load data from selected local month folders."""
    selected_months = st.sidebar.multiselect(
        "Choose Month(s) to Load",
        options=list(folders.keys()),
        default=list(folders.keys())
    )
    if not selected_months:
        st.sidebar.warning("Please select at least one month folder.")
        return None, ""

    month_data_dict = {}
    validation_messages = []
    for m_choice in selected_months:
        m_dir = folders[m_choice]
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
            cached_data = cached_reconcile_month(dtb_path, is_path, cons_path, m_label)
            month_data_dict[m_label] = cached_data
    if validation_messages:
        show_validation_banner(validation_messages)
    if month_data_dict:
        return combine_month_datasets(month_data_dict)
    st.sidebar.warning("No valid month data could be loaded from the selected folders.")
    return None, ""


def _load_uploaded_files() -> tuple:
    """Load data from uploaded Excel files (single or multiple months)."""
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

        for uploaded_file in [dtb_file, is_file, cons_file]:
            if uploaded_file and (".." in uploaded_file.name or "/" in uploaded_file.name or "\\" in uploaded_file.name):
                validation_messages.append(f"Invalid filename detected: {uploaded_file.name}. Filenames cannot contain path separators or '..'.")

        if validation_messages:
            show_validation_banner(validation_messages, level="error")
        else:
            with st.spinner(f"Processing {selected_month_label}..."):
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
                    if is_path:
                        os.unlink(is_path)
                    if cons_path:
                        os.unlink(cons_path)
                except Exception:
                    pass

                st.toast(f"Added {selected_month_label} successfully!")
                st.rerun()

    if st.session_state.uploaded_months:
        return combine_month_datasets(st.session_state.uploaded_months)
    return None, ""


def _load_zip_months() -> tuple:
    """Load data from a ZIP archive of month folders."""
    zip_file = st.sidebar.file_uploader("Upload ZIP of month folders", type=["zip"])
    if not zip_file:
        return None, ""

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
            return combine_month_datasets(month_data)
    return None, ""


def _render_executive_dashboard(cat_df, top_items):
    """Render Executive Dashboard tab."""
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


def _render_item_consumption(trx_df, item_df):
    """Render Item Consumption & Repeated Orders tab."""
    st.subheader("📦 Item Consumption & Repeated Orders Summary")
    st.markdown(
        "Track monthly cumulative quantities, average unit prices, and total spend per item "
        "across recurring orders (e.g. **Cleo 330ml mineral water**, **Toilet Tissue**, **Slippers**, **Soap**, **Garbage Bags**)."
    )

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

        st.markdown("#### 🔬 Item Inspector: Delivery & Order History")
        item_options = filtered_items['Item_Name'].unique().tolist()
        default_idx = next((i for i, name in enumerate(item_options) if "cleo" in name.lower()), 0)
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

            st.markdown(f"**Individual Order Vouchers for {selected_item_name}:**")
            voucher_view = item_trxs[['Date', 'Voucher_Ref', 'Qty', 'Unit', 'Unit_Price', 'Amount', 'JRNL', 'Raw_Description']].copy()
            voucher_view['Unit_Price'] = voucher_view['Unit_Price'].apply(format_idr)
            voucher_view['Amount'] = voucher_view['Amount'].apply(format_idr)
            st.dataframe(voucher_view, width="stretch")

        item_csv = filtered_items.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Item Totals Summary (CSV)",
            data=item_csv,
            file_name=f"Item_Totals_{sel_item_cat.replace(' ', '_')}.csv",
            mime="text/csv"
        )
    else:
        st.info("No items match the selected criteria.")


def _render_transaction_drilldown(trx_df):
    """Render Transaction Drilldown tab."""
    st.subheader("Detailed Spending Explorer")
    st.markdown("Inspect every individual transaction, vendor voucher, item quantity, and line amount.")

    fcol1, fcol2, fcol3 = st.columns([2, 2, 2])
    with fcol1:
        all_cats = ["All Categories"] + sorted(trx_df['Category'].unique().tolist())
        selected_cat = st.selectbox("Filter by Category", all_cats)
    with fcol2:
        search_query = st.text_input("🔍 Search Item / Vendor / Voucher", "")
    with fcol3:
        sort_by = st.selectbox("Sort By", ["Amount (High to Low)", "Amount (Low to High)", "Date (Latest)", "Item Name"])

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

    if sort_by == "Amount (High to Low)":
        filtered_df = filtered_df.sort_values(by="Amount", ascending=False)
    elif sort_by == "Amount (Low to High)":
        filtered_df = filtered_df.sort_values(by="Amount", ascending=True)
    elif sort_by == "Date (Latest)":
        filtered_df = filtered_df.sort_values(by="Date", ascending=False)
    elif sort_by == "Item Name":
        filtered_df = filtered_df.sort_values(by="Item_Name", ascending=True)

    st.caption(f"Showing {len(filtered_df)} transactions totaling **{format_idr(filtered_df['Amount'].sum())}**")

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

    csv_slice = filtered_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Filtered Results (CSV)",
        data=csv_slice,
        file_name=f"Filtered_Expenses_{selected_cat.replace(' ', '_')}.csv",
        mime="text/csv"
    )


def _render_budget_variance(cat_df):
    """Render Budget vs Actual Variance tab."""
    st.subheader("Budget vs. Actual Variance Analysis")
    st.markdown("Identify budget overruns and operational savings at a glance.")

    over_bgt = cat_df[cat_df['Status'] == 'Over Budget']
    if not over_bgt.empty:
        st.warning(f"⚠️ **Attention:** {len(over_bgt)} categories exceeded their budget allocation this month!")
        cols = st.columns(min(len(over_bgt), 3))
        for i, (_, row) in enumerate(over_bgt.iterrows()):
            with cols[i % len(cols)]:
                st.error(f"**{row['Category']}**\n\nActual: {format_idr(row['Actual'])}\n\nBudget: {format_idr(row['Budget'])}\n\n**Over by {format_idr(row['Variance_IDR'])} (+{row['Variance_Pct']:.1f}%)**")

    st.markdown("<br>", unsafe_allow_html=True)

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


def _render_mom_comparison(folders: dict, selected_month_name: str, reconciled_data: dict):
    """Render Month-over-Month Comparison tab."""
    st.subheader("Month-over-Month Spending Comparison")
    mom_folder_pairs = list(folders.items())
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


def _render_export(reconciled_data: dict, selected_month_name: str, trx_df: pd.DataFrame) -> bytes:
    """Render Separated Excel & Export tab; return generated Excel bytes."""
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

    excel_bytes = exporter.create_separated_excel(reconciled_data, selected_month_name)

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

    return excel_bytes


def render_monthly_tracker() -> None:
    """Render the Monthly Expense Tracker UI."""
    base_dir, has_local, folders = _get_month_folders()
    base_dir, has_local, folders = _sidebar_data_dir(base_dir, has_local, folders)

    if st.session_state.get("demo_mode"):
        data_source = "Demo Data"
    else:
        data_source = st.sidebar.radio(
            "Select Data Source",
            ["Select Month Folder", "Upload Custom Excel Files", "Upload Multiple Months"],
            index=0 if has_local else 1
        )

    reconciled_data = None
    selected_month_name = "Demo August 2026"

    if data_source == "Demo Data":
        reconciled_data, selected_month_name = _load_demo_data()
    elif data_source == "Select Month Folder":
        result = _load_selected_months(folders)
        reconciled_data, selected_month_name = result if result[0] is not None else (None, "")
    elif data_source == "Upload Custom Excel Files":
        result = _load_uploaded_files()
        reconciled_data, selected_month_name = result if result[0] is not None else (None, "")
    elif data_source == "Upload Multiple Months":
        result = _load_zip_months()
        reconciled_data, selected_month_name = result if result[0] is not None else (None, "")

    if not reconciled_data:
        st.info("Please select or upload Excel files from the sidebar to begin.")
        return

    metrics = reconciled_data['metrics']
    cat_df = reconciled_data['category_summary']
    trx_df = reconciled_data['transactions']
    top_items = reconciled_data['top_cost_drivers']
    item_df = reconciled_data.get('item_summary', pd.DataFrame())

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

    st.markdown('<div class="main-header">Hotel Expense Separator & Spending Tracker</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Hotel • Room Division & Housekeeping • Period: <b>{selected_month_name}</b></div>', unsafe_allow_html=True)

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

    tab_dash, tab_items, tab_drill, tab_bgt, tab_mom, tab_export = st.tabs([
        "📊 Executive Dashboard",
        "📦 Item Totals & Quantities",
        "🔍 Transaction Drilldown",
        "⚖️ Budget vs Actual Variance",
        "📈 Month-over-Month Comparison",
        "📂 Separated Excel & Export"
    ])

    with tab_dash:
        _render_executive_dashboard(cat_df, top_items)

    with tab_items:
        _render_item_consumption(trx_df, item_df)

    with tab_drill:
        _render_transaction_drilldown(trx_df)

    with tab_bgt:
        _render_budget_variance(cat_df)

    with tab_mom:
        _render_mom_comparison(folders, selected_month_name, reconciled_data)

    with tab_export:
        _render_export(reconciled_data, selected_month_name, trx_df)
