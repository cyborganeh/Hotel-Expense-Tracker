"""Warehouse Stock Request (SR) Separator view."""
import glob
import hashlib
import os
import re

import pandas as pd
import streamlit as st

import sr_parser
import parser
import exporter
import ui_components
from ui.formatting import format_idr, show_validation_banner
from ui.security import sanitize_html, validate_upload_size, safe_filename
from data.months import get_default_data_dir


def render_sr_separator() -> None:
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
        sr_options.append(f"Load Local {parser.infer_month_label(os.path.basename(BASE_SR_DIR))} SRs")
    sr_options.append("Upload SR PDFs")

    sr_source = st.sidebar.radio("Select Source", sr_options, index=0)

    sr_df = pd.DataFrame()
    local_sr_option = f"Load Local {parser.infer_month_label(os.path.basename(BASE_SR_DIR))} SRs"
    sr_report_title = parser.infer_month_label(os.path.basename(BASE_SR_DIR))

    if sr_source == local_sr_option:
        st.sidebar.markdown("**Local Folder:** `SR SEPTEMBER`")
        local_pdfs = sorted(glob.glob(os.path.join(BASE_SR_DIR, "*.pdf")))
        st.sidebar.caption(f"Found {len(local_pdfs)} PDF files:")
        for lp in local_pdfs:
            st.sidebar.caption(f"📄 {os.path.basename(lp)}")
        if local_pdfs:
            with st.spinner("Parsing local SR PDF files..."):
                sr_df = sr_parser.parse_multiple_sr_pdfs(local_pdfs)
            sr_report_title = parser.infer_month_label(os.path.basename(BASE_SR_DIR))

    elif sr_source == "Upload SR PDFs":
        uploaded_pdfs = st.sidebar.file_uploader(
            "Upload Stock Request PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            help="Upload one or multiple SR PDF files from Gudang Central."
        )
        sr_label = st.sidebar.text_input("Report Title / Month", "Stock Request Report")
        sr_report_title = sanitize_html(sr_label)
        if uploaded_pdfs:
            for uploaded in uploaded_pdfs:
                validate_upload_size(uploaded)
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

        **Supported formats**: Stock Request Consumption PDFs from Hotel (`.pdf`).
        """)
        return

    # Main Page UI
    st.markdown('<div class="main-header">Warehouse Stock Request (SR) Separator</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Hotel • Gudang Central Issuing to Housekeeping • <b>{sr_report_title}</b></div>', unsafe_allow_html=True)

    identity_columns = [column for column in ['Date', 'SR_Number', 'Item_Code', 'Qty', 'Total'] if column in sr_df]
    period_identity = sr_report_title + "\x1f" + sr_df[identity_columns].to_csv(index=False)
    period_key = hashlib.sha256(period_identity.encode("utf-8")).hexdigest()[:12]
    input_col_occupancy, input_col_rooms_sold = st.columns(2)
    with input_col_occupancy:
        occupancy_pct = st.number_input(
            "Occupancy (%)", min_value=0.0, max_value=100.0, value=0.0,
            step=0.1, format="%.1f", key=f"sr_occupancy_{period_key}"
        )
    with input_col_rooms_sold:
        total_rooms_sold = st.number_input(
            "Total Room Sold", min_value=0, value=0, step=1,
            key=f"sr_rooms_sold_{period_key}"
        )

    # Top KPI Metrics
    tot_amt = sr_df['Total'].sum()
    tot_qty = sr_df['Qty'].sum()
    tot_items = sr_df['Item_Name'].nunique()
    tot_order = sr_df['SR_Number'].nunique()
    rooms_sold_valid = total_rooms_sold > 0
    spend_per_room_sold = tot_amt / total_rooms_sold if rooms_sold_valid else None

    st.markdown("#### Period Performance")
    ui_components.render_kpi_row([
        {'label': 'OCCUPANCY', 'value': f"{occupancy_pct:.1f}%",
         'delta': 'Period input', 'delta_tone': 'neutral', 'show_arrow': False},
        {'label': 'ROOMS SOLD', 'value': f"{total_rooms_sold:,.0f}",
         'delta': 'Period total', 'delta_tone': 'neutral', 'show_arrow': False},
        {'label': 'SR SPEND / ROOM SOLD',
         'value': format_idr(spend_per_room_sold) if rooms_sold_valid else 'N/A',
         'delta': 'Total SR spend per room sold', 'delta_tone': 'neutral', 'show_arrow': False},
    ])

    excel_sr_bytes = exporter.create_sr_separated_excel(
        sr_df,
        sr_report_title,
        occupancy_pct=occupancy_pct,
        total_rooms_sold=total_rooms_sold,
    )
    st.sidebar.divider()
    st.sidebar.markdown("### 📥 Quick Export")
    safe_title = safe_filename(sr_report_title, default="SR_Report")
    st.sidebar.download_button(
        label="Download Separated SR Excel (.xlsx)",
        data=excel_sr_bytes,
        file_name=f"Separated_SR_{safe_title.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        type="primary"
    )

    ui_components.render_kpi_row([
        {'label': 'TOTAL SPEND (SR)', 'value': format_idr(tot_amt),
         'delta': f"{len(sr_df)} line items issued", 'delta_tone': 'neutral', 'show_arrow': True},
        {'label': 'TOTAL QUANTITY', 'value': f"{tot_qty:,.0f}",
         'delta': 'Across all units & categories', 'delta_tone': 'pos', 'show_arrow': True},
        {'label': 'UNIQUE ITEMS', 'value': str(tot_items),
         'delta': 'Distinct products ordered', 'delta_tone': 'neutral', 'show_arrow': True},
        {'label': 'SR ORDERS COUNT', 'value': str(tot_order),
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
        cat_sum['Spend_Per_Room_Sold'] = (
            cat_sum['Total_Amount'].div(total_rooms_sold).apply(format_idr)
            if rooms_sold_valid else 'N/A'
        )

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
        disp_cat['Spend_Per_Room_Sold'] = disp_cat['Spend_Per_Room_Sold'].astype(str)
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
        cat_spend_per_room_sold = cat_amt / total_rooms_sold if rooms_sold_valid else None

        st.markdown(f"### {emoji} {category_name}")
        cat_spend_per_room_text = format_idr(cat_spend_per_room_sold) if rooms_sold_valid else 'N/A'
        st.markdown(f"**Total Spend:** {format_idr(cat_amt)} ({pct_all:.1f}% of SR spend) • **Spend / Room Sold:** {cat_spend_per_room_text} • **Total Qty:** {cat_q:,.0f} • **Unique Items:** {len(c_items)}")

        st.markdown("#### Item Totals Summary (Repeated Orders Aggregated)")
        disp_items = c_items.copy()
        disp_items['Qty_Per_Room_Sold'] = (
            disp_items['Total_Qty'] / total_rooms_sold if rooms_sold_valid else 'N/A'
        )
        disp_items['Total_Qty'] = disp_items['Total_Qty'].apply(lambda x: f"{x:,.0f}")
        disp_items['Qty_Per_Room_Sold'] = disp_items['Qty_Per_Room_Sold'].apply(
            lambda x: f"{x:,.4f}" if rooms_sold_valid else x
        )
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
        - 📊 **SR Summary**: Period occupancy, room sold, spend per room sold, category breakdown, and top 10 cost drivers.
        - 🛎️ **Guest Supplies**: Item quantities per room sold + chronological issuing history.
        - 🧹 **Cleaning Supplies**: Item quantities per room sold + order log.
        - 🧻 **Paper Supplies**: Item quantities per room sold + order log.
        - 📑 **Print & Stationery**: Item quantities per room sold + order log.
        - 📋 **All SR Items**: Complete line item audit table.
        """)

        safe_title = safe_filename(sr_report_title, default="SR_Report")
        st.download_button(
            label=f"📥 Download Full Separated SR Excel ({sr_report_title})",
            data=excel_sr_bytes,
            file_name=f"Separated_SR_{safe_title.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
