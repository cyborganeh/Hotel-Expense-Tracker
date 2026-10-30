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
import room_cost


def _format_room_cost_value(value) -> str:
    return "Belum lengkap" if pd.isna(value) else format_idr(value)


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

    st.markdown('<div class="main-header">Warehouse Stock Request (SR) Separator</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Hotel • Gudang Central Issuing to Housekeeping • <b>{sr_report_title}</b></div>', unsafe_allow_html=True)

    if sr_df.empty:
        ui_components.render_empty_state(
            "No SR PDFs loaded. Daily room costs can still be entered; the SR component is currently Rp 0."
        )
        st.markdown("""
        ### About Warehouse Stock Request (SR) Separator
        This tool extracts items from Gudang Central Stock Request PDFs and separates them automatically into:
        - 🛎️ **Guest Supplies**: Soap, slipper, toothbrush, shower cap, coffee, tea, sugar, Cleo water, etc.
        - 🧹 **Cleaning Supplies**: Garbage bags (Plastik sampah), cleaning chemicals, glass cleaner, etc.
        - 🧻 **Paper Supplies**: Facial tissue, hand towel, toilet roll, plastic roll.
        - 📑 **Print & Stationery**: Form blanks, staplers, tape, stationery items.

        **Supported formats**: Stock Request Consumption PDFs from Hotel (`.pdf`).
        """)
        sr_df = pd.DataFrame(columns=[
            "Date", "SR_Number", "Requested_By", "Department", "Item_Code", "Item_Name",
            "Category", "Qty", "Unit", "Cost", "Total", "Source_File",
        ])

    identity_columns = [column for column in ['Date', 'SR_Number', 'Item_Code', 'Qty', 'Total'] if column in sr_df]
    period_identity = sr_report_title + "\x1f" + sr_df[identity_columns].to_csv(index=False)
    period_key = hashlib.sha256(period_identity.encode("utf-8")).hexdigest()[:12]

    current_month = pd.Timestamp.today().strftime("%Y-%m")
    sr_dates = pd.to_datetime(sr_df.get("Date", pd.Series(dtype=str)), errors="coerce", format="mixed")
    available_months = {current_month}
    available_months.update(sr_dates.dropna().dt.strftime("%Y-%m").tolist())
    room_cost_dir = room_cost.get_room_cost_storage_dir()
    if os.path.isdir(room_cost_dir):
        available_months.update(
            name[:-5] for name in os.listdir(room_cost_dir)
            if re.fullmatch(r"\d{4}-\d{2}\.json", name)
        )
    month_options = sorted(available_months, reverse=True)
    preferred_month = next(
        (month for month in month_options if month in set(sr_dates.dropna().dt.strftime("%Y-%m"))),
        current_month,
    )
    selected_month = st.sidebar.selectbox(
        "Room Cost Month (YYYY-MM)",
        month_options,
        index=month_options.index(preferred_month),
        key="sr_room_cost_month",
    )
    try:
        saved_daily_costs = room_cost.load_daily_room_costs(selected_month)
    except ValueError as exc:
        saved_daily_costs = pd.DataFrame(columns=room_cost.DAILY_COLUMNS)
        show_validation_banner([str(exc)], level="error")

    st.sidebar.markdown("### Daily Room Cost Excel")
    uploaded_cost_excel = st.sidebar.file_uploader(
        "Upload daily costs (.xlsx)",
        type=["xlsx"],
        key=f"room_cost_upload_{selected_month}",
        help="Required columns: Date, Room Nights, Water, Electricity, and Linen Laundry.",
    )
    if uploaded_cost_excel is not None:
        try:
            validate_upload_size(uploaded_cost_excel)
            imported_daily_costs = room_cost.parse_daily_room_cost_excel(uploaded_cost_excel, selected_month)
            duplicate_dates = sorted(set(saved_daily_costs["Date"]) & set(imported_daily_costs["Date"]))
            if duplicate_dates:
                st.sidebar.warning("Data exists for: " + ", ".join(duplicate_dates))
                replace_confirmed = st.sidebar.checkbox(
                    "Confirm replacement for these dates",
                    key=f"room_cost_replace_confirm_{selected_month}",
                )
            else:
                replace_confirmed = True

            if st.sidebar.button(
                "Save Daily Costs",
                key=f"room_cost_save_{selected_month}",
                disabled=bool(duplicate_dates) and not replace_confirmed,
                type="primary",
            ):
                try:
                    merged_daily_costs = room_cost.merge_daily_room_costs(
                        saved_daily_costs,
                        imported_daily_costs,
                        replace_dates=replace_confirmed,
                        month=selected_month,
                    )
                    room_cost.save_daily_room_costs(selected_month, merged_daily_costs)
                    st.sidebar.success(f"Saved {len(imported_daily_costs)} daily row(s).")
                    st.rerun()
                except Exception as exc:
                    show_validation_banner([f"Could not save daily room-cost data: {exc}"], level="error")
            st.sidebar.caption(f"Validated {len(imported_daily_costs)} row(s) for {selected_month}.")
        except Exception as exc:
            show_validation_banner([f"Daily room-cost Excel: {exc}"], level="error")
    else:
        st.sidebar.caption(f"{len(saved_daily_costs)} saved daily row(s) for {selected_month}.")

    st.sidebar.markdown("### Night Audit Occupancy & Room Sold")
    uploaded_night_audit = st.sidebar.file_uploader(
        "Upload Night Audit Report (.xlsx)",
        type=["xlsx"],
        key=f"night_audit_upload_{selected_month}",
        help="Extracts occupancy and room sold by date. Existing cost values are preserved.",
    )
    if uploaded_night_audit is not None:
        try:
            validate_upload_size(uploaded_night_audit)
            audit_rows = room_cost.parse_night_audit_report(uploaded_night_audit, selected_month)
            duplicate_audit_dates = sorted(set(saved_daily_costs["Date"]) & set(audit_rows["Date"]))
            if duplicate_audit_dates:
                st.sidebar.warning("Room-night data exists for: " + ", ".join(duplicate_audit_dates))
                replace_audit_confirmed = st.sidebar.checkbox(
                    "Confirm Night Audit replacement",
                    key=f"night_audit_replace_confirm_{selected_month}",
                )
            else:
                replace_audit_confirmed = True

            if st.sidebar.button(
                "Save Occupancy & Room Sold",
                key=f"night_audit_save_{selected_month}",
                disabled=bool(duplicate_audit_dates) and not replace_audit_confirmed,
                type="primary",
            ):
                merged_audit_rows = room_cost.merge_daily_room_costs(
                    saved_daily_costs,
                    audit_rows,
                    replace_dates=replace_audit_confirmed,
                    month=selected_month,
                )
                room_cost.save_daily_room_costs(selected_month, merged_audit_rows)
                st.sidebar.success(
                    f"Saved {len(audit_rows)} dates; room sold total {audit_rows['Room_Nights'].sum():,.0f}."
                )
                st.rerun()
            st.sidebar.caption(
                f"Detected {len(audit_rows)} dates and {audit_rows['Room_Nights'].sum():,.0f} room sold."
            )
        except Exception as exc:
            show_validation_banner([f"Night Audit report: {exc}"], level="error")

    st.sidebar.markdown("### Energy Cost")
    uploaded_energy_cost = st.sidebar.file_uploader(
        "Upload Energy Cost workbook (.xlsx)",
        type=["xlsx"],
        key=f"energy_cost_upload_{selected_month}",
        help="Extracts daily electricity and water amounts; existing room sold and linen data are preserved.",
    )
    if uploaded_energy_cost is not None:
        try:
            validate_upload_size(uploaded_energy_cost)
            energy_rows = room_cost.parse_energy_cost_workbook(uploaded_energy_cost, selected_month)
            duplicate_energy_dates = sorted(set(saved_daily_costs["Date"]) & set(energy_rows["Date"]))
            if duplicate_energy_dates:
                st.sidebar.warning("Energy data exists for: " + ", ".join(duplicate_energy_dates))
                replace_energy_confirmed = st.sidebar.checkbox(
                    "Confirm energy value replacement",
                    key=f"energy_cost_replace_confirm_{selected_month}",
                )
            else:
                replace_energy_confirmed = True

            if st.sidebar.button(
                "Save Energy Costs",
                key=f"energy_cost_save_{selected_month}",
                disabled=bool(duplicate_energy_dates) and not replace_energy_confirmed,
                type="primary",
            ):
                merged_energy_rows = room_cost.merge_daily_room_costs(
                    saved_daily_costs,
                    energy_rows,
                    replace_dates=replace_energy_confirmed,
                    month=selected_month,
                )
                room_cost.save_daily_room_costs(selected_month, merged_energy_rows)
                st.sidebar.success(
                    f"Saved electricity Rp {energy_rows['Electricity'].sum():,.0f} and water Rp {energy_rows['Water'].sum():,.0f}."
                )
                st.rerun()
            st.sidebar.caption(f"Detected energy data for {len(energy_rows)} date(s).")
        except Exception as exc:
            show_validation_banner([f"Energy Cost workbook: {exc}"], level="error")

    st.sidebar.markdown("### Daily Linen Cost")
    uploaded_linen_cost = st.sidebar.file_uploader(
        "Upload Daily Linen Cost workbook (.xlsx)",
        type=["xlsx"],
        key=f"linen_cost_upload_{selected_month}",
        help="Reads TOTAL BIAYA HARIAN by date and preserves saved room, water, and electricity data.",
    )
    if uploaded_linen_cost is not None:
        try:
            validate_upload_size(uploaded_linen_cost)
            linen_rows = room_cost.parse_daily_linen_cost_workbook(uploaded_linen_cost, selected_month)
            duplicate_linen_dates = sorted(set(saved_daily_costs["Date"]) & set(linen_rows["Date"]))
            if duplicate_linen_dates:
                st.sidebar.warning("Linen data exists for: " + ", ".join(duplicate_linen_dates))
                replace_linen_confirmed = st.sidebar.checkbox(
                    "Confirm linen value replacement",
                    key=f"linen_cost_replace_confirm_{selected_month}",
                )
            else:
                replace_linen_confirmed = True

            if st.sidebar.button(
                "Save Linen Costs",
                key=f"linen_cost_save_{selected_month}",
                disabled=bool(duplicate_linen_dates) and not replace_linen_confirmed,
                type="primary",
            ):
                merged_linen_rows = room_cost.merge_daily_room_costs(
                    saved_daily_costs,
                    linen_rows,
                    replace_dates=replace_linen_confirmed,
                    month=selected_month,
                )
                room_cost.save_daily_room_costs(selected_month, merged_linen_rows)
                st.sidebar.success(f"Saved linen total Rp {linen_rows['Linen_Laundry'].sum():,.0f}.")
                st.rerun()
            st.sidebar.caption(f"Detected linen costs for {len(linen_rows)} date(s).")
        except Exception as exc:
            show_validation_banner([f"Daily Linen Cost workbook: {exc}"], level="error")

    room_cost_daily_df = pd.DataFrame()
    room_cost_monthly_df = pd.DataFrame()
    room_cost_warnings = []
    if not saved_daily_costs.empty:
        try:
            room_cost_daily_df, room_cost_monthly_df, room_cost_warnings = room_cost.calculate_daily_room_costs(
                sr_df, saved_daily_costs, selected_month
            )
            for warning in room_cost_warnings:
                st.warning(warning)
        except ValueError as exc:
            show_validation_banner([f"Room-cost calculation: {exc}"], level="error")

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
        room_cost_daily_df=room_cost_daily_df,
        room_cost_monthly_df=room_cost_monthly_df,
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
    st.sidebar.download_button(
        label="Download Morning Briefing Template (.xlsx)",
        data=exporter.create_morning_briefing_template(
            selected_month, sr_data=sr_df, daily_cost_data=saved_daily_costs
        ),
        file_name=f"Morning_Briefing_{selected_month}_with_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
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
    tab_overview, tab_gs, tab_cs, tab_ps, tab_pst, tab_all, tab_exp, tab_room_cost = st.tabs([
        "📊 Category Overview",
        "🛎️ Guest Supplies",
        "🧹 Cleaning Supplies",
        "🧻 Paper Supplies",
        "📑 Print & Stationery",
        "📋 All Stock Requests Log",
        "📥 Excel Export",
        "🏨 Daily Room Cost"
    ])

    cat_order = ['Guest Supplies', 'Cleaning Supplies', 'Paper Supplies', 'Print & Stationery']

    # Tab 1: Category Overview
    with tab_overview:
        col_ch1, col_ch2 = st.columns([1, 1])

        cat_sum = sr_parser.get_category_summary(sr_df)
        if cat_sum.empty:
            ui_components.render_empty_state("SR category summaries will appear after SR PDFs are loaded.")
        else:
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
            ui_components.render_empty_state(f"No {category_name} items recorded in the uploaded SR files.")
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

    with tab_room_cost:
        st.markdown(f"### Daily Room Cost • {selected_month}")
        month_start_date = pd.Timestamp(f"{selected_month}-01").date()
        month_end_date = pd.Period(selected_month).end_time.date()
        default_entry_date = (
            min(pd.Timestamp.today().date() - pd.Timedelta(days=1), month_end_date)
            if selected_month == pd.Timestamp.today().strftime("%Y-%m")
            else month_end_date
        )
        manual_date = st.date_input(
            "Manual entry date",
            value=default_entry_date,
            min_value=month_start_date,
            max_value=month_end_date,
            key=f"manual_entry_date_{selected_month}",
        )
        manual_date_text = manual_date.isoformat()
        selected_daily_row = saved_daily_costs.loc[saved_daily_costs["Date"] == manual_date_text]
        selected_daily_row = selected_daily_row.iloc[0] if not selected_daily_row.empty else None
        manual_data = room_cost.load_manual_briefing_data(manual_date_text)
        manual_entries = manual_data["Entries"]

        def _daily_default(column: str, default: float = 0.0) -> float:
            if selected_daily_row is None or pd.isna(selected_daily_row.get(column)):
                return default
            return float(selected_daily_row[column])

        def _editor_values(entry_type: str, columns: list[str]) -> pd.DataFrame:
            subset = manual_entries.loc[manual_entries["Entry_Type"] == entry_type, columns]
            subset = subset.reset_index(drop=True).copy()
            for numeric_column in ("Room_Credit", "Qty", "Unit_Cost"):
                if numeric_column in subset.columns:
                    subset[numeric_column] = pd.to_numeric(subset[numeric_column], errors="coerce")
            return subset

        st.markdown("#### Input Manual Harian")
        with st.form(f"manual_daily_entry_form_{selected_month}_{manual_date_text}"):
            st.caption("Isi Room Cost, staff, credit RA, laundry, aktivitas area, dan follow-up untuk tanggal yang dipilih.")
            cost_col1, cost_col2, cost_col3 = st.columns(3)
            with cost_col1:
                manual_occupancy = st.number_input(
                    "Occupancy (%)", min_value=0.0, max_value=100.0,
                    value=_daily_default("Occupancy"), step=0.1, format="%.2f",
                )
                manual_room_nights = st.number_input(
                    "Room Sold", min_value=0, step=1,
                    value=int(_daily_default("Room_Nights")),
                )
            with cost_col2:
                manual_water = st.number_input(
                    "Air (IDR)", min_value=0.0, value=_daily_default("Water"), step=1000.0,
                )
                manual_electricity = st.number_input(
                    "Listrik (IDR)", min_value=0.0, value=_daily_default("Electricity"), step=1000.0,
                )
            with cost_col3:
                manual_linen = st.number_input(
                    "Linen (IDR)", min_value=0.0, value=_daily_default("Linen_Laundry"), step=1000.0,
                )
                manual_laundry_cost = st.number_input(
                    "Bon Vivo Guest Laundry (IDR)", min_value=0.0,
                    value=float(manual_entries.loc[manual_entries["Entry_Type"] == "guest_laundry", "Total"].iloc[-1])
                    if not manual_entries.loc[manual_entries["Entry_Type"] == "guest_laundry"].empty
                    and pd.notna(manual_entries.loc[manual_entries["Entry_Type"] == "guest_laundry", "Total"].iloc[-1]) else 0.0,
                    step=1000.0,
                )
            laundry_room_nights_default = int(
                manual_entries.loc[manual_entries["Entry_Type"] == "guest_laundry", "Qty"].iloc[-1]
            ) if not manual_entries.loc[manual_entries["Entry_Type"] == "guest_laundry"].empty and pd.notna(
                manual_entries.loc[manual_entries["Entry_Type"] == "guest_laundry", "Qty"].iloc[-1]
            ) else 0
            laundry_room_nights = st.number_input(
                "Guest Laundry Room Nights", min_value=0, step=1, value=laundry_room_nights_default,
            )

            staff_incharge = st.text_input(
                "Staff Incharge", value=manual_data["Staff_Incharge"],
            )
            remarks = st.text_area("Remark / Follow Up", value=manual_data["Remarks"])

            st.markdown("**Room Attendant / Room Credit**")
            ra_editor_columns = ["Staff_Name", "Section", "Room_Credit", "Description", "Notes"]
            ra_editor = st.data_editor(
                _editor_values("ra", ra_editor_columns), num_rows="dynamic", hide_index=True,
                column_config={
                    "Staff_Name": st.column_config.TextColumn("Room Attendant"),
                    "Section": st.column_config.TextColumn("Section / Floor"),
                    "Room_Credit": st.column_config.NumberColumn("Room Credit", min_value=0, step=1),
                    "Description": st.column_config.TextColumn("Activity / Project"),
                    "Notes": st.column_config.TextColumn("Notes"),
                },
                key=f"manual_ra_editor_{selected_month}_{manual_date_text}",
            )
            st.markdown("**Project Public Area**")
            project_editor_columns = ["Area", "Description", "Qty", "Unit", "Unit_Cost", "Status", "PIC", "Notes"]
            project_editor = st.data_editor(
                _editor_values("project_pa", project_editor_columns), num_rows="dynamic", hide_index=True,
                column_config={
                    "Area": st.column_config.TextColumn("Project / Area"),
                    "Description": st.column_config.TextColumn("Description"),
                    "Qty": st.column_config.NumberColumn("Qty", min_value=0, step=1),
                    "Unit": st.column_config.TextColumn("Unit"),
                    "Unit_Cost": st.column_config.NumberColumn("Unit Cost (IDR)", min_value=0, step=1000),
                    "Status": st.column_config.TextColumn("Status"),
                    "PIC": st.column_config.TextColumn("PIC"),
                    "Notes": st.column_config.TextColumn("Notes"),
                },
                key=f"manual_project_editor_{selected_month}_{manual_date_text}",
            )
            st.markdown("**Daily Public Area**")
            daily_pa_editor_columns = ["Area", "Description", "Qty", "Unit_Cost", "PIC", "Notes"]
            daily_pa_editor = st.data_editor(
                _editor_values("daily_pa", daily_pa_editor_columns), num_rows="dynamic", hide_index=True,
                column_config={
                    "Area": st.column_config.TextColumn("Area"),
                    "Description": st.column_config.TextColumn("Description"),
                    "Qty": st.column_config.NumberColumn("Qty", min_value=0, step=1),
                    "Unit_Cost": st.column_config.NumberColumn("Unit Cost (IDR)", min_value=0, step=1000),
                    "PIC": st.column_config.TextColumn("PIC"),
                    "Notes": st.column_config.TextColumn("Notes"),
                },
                key=f"manual_daily_pa_editor_{selected_month}_{manual_date_text}",
            )
            save_manual_entry = st.form_submit_button("Simpan Input Manual", type="primary")

        if save_manual_entry:
            if manual_room_nights <= 0:
                st.error("Room Sold harus lebih dari 0 untuk menyimpan biaya per kamar.")
            else:
                manual_cost_row = pd.DataFrame([{
                    "Date": manual_date_text,
                    "Occupancy": manual_occupancy,
                    "Room_Nights": manual_room_nights,
                    "Water": manual_water,
                    "Electricity": manual_electricity,
                    "Linen_Laundry": manual_linen,
                }])
                merged_manual_costs = room_cost.merge_daily_room_costs(
                    saved_daily_costs, manual_cost_row, replace_dates=True, month=selected_month,
                )
                room_cost.save_daily_room_costs(selected_month, merged_manual_costs)

                manual_records = []
                editor_mappings = [
                    ("ra", ra_editor, {
                        "Staff_Name": "Staff_Name", "Section": "Section", "Room_Credit": "Room_Credit",
                        "Description": "Description", "Notes": "Notes",
                    }),
                    ("project_pa", project_editor, {
                        "Area": "Area", "Description": "Description", "Qty": "Qty", "Unit": "Unit",
                        "Unit_Cost": "Unit_Cost", "Status": "Status", "PIC": "PIC", "Notes": "Notes",
                    }),
                    ("daily_pa", daily_pa_editor, {
                        "Area": "Area", "Description": "Description", "Qty": "Qty",
                        "Unit_Cost": "Unit_Cost", "PIC": "PIC", "Notes": "Notes",
                    }),
                ]
                for entry_type, frame, mapping in editor_mappings:
                    for editor_row in frame.to_dict(orient="records"):
                        if not any(value is not None and not pd.isna(value) and str(value).strip() for value in editor_row.values()):
                            continue
                        record = {"Entry_Type": entry_type}
                        record.update({target: editor_row.get(source) for source, target in mapping.items()})
                        manual_records.append(record)
                if manual_laundry_cost or laundry_room_nights:
                    manual_records.append({
                        "Entry_Type": "guest_laundry",
                        "Qty": laundry_room_nights,
                        "Total": manual_laundry_cost,
                    })
                room_cost.save_manual_briefing_data(
                    manual_date_text,
                    staff_incharge=staff_incharge,
                    remarks=remarks,
                    entries=pd.DataFrame(manual_records),
                )
                st.success(f"Input manual tanggal {manual_date_text} tersimpan di SQLite.")
                st.rerun()

        if room_cost_daily_df.empty:
            ui_components.render_empty_state("Upload and save the daily cost Excel from the sidebar to calculate room cost.")
        else:
            latest = room_cost_monthly_df.iloc[0]
            latest_daily = room_cost_daily_df.iloc[-1]
            ui_components.render_kpi_row([
                {'label': 'ROOM NIGHTS MTD', 'value': f"{latest['Room_Nights']:,.0f}",
                 'delta': selected_month, 'delta_tone': 'neutral', 'show_arrow': False},
                {'label': 'SR COST / ROOM', 'value': format_idr(latest['SR_Per_Room']),
                 'delta': 'Current MTD rate, applied retroactively', 'delta_tone': 'neutral', 'show_arrow': False},
                {'label': 'LATEST DAILY COST / ROOM', 'value': _format_room_cost_value(latest_daily['Total_Cost_Per_Room']),
                 'delta': str(latest_daily['Date']), 'delta_tone': 'neutral', 'show_arrow': False},
                {'label': 'TOTAL COST / ROOM', 'value': format_idr(latest['Total_Cost_Per_Room']),
                 'delta': 'Utilities, linen, and SR', 'delta_tone': 'neutral', 'show_arrow': False},
                {'label': 'TOTAL ROOM COST', 'value': format_idr(latest['Total_Room_Cost']),
                 'delta': 'Current MTD total', 'delta_tone': 'neutral', 'show_arrow': False},
            ])

            st.markdown("#### Monthly-to-date Summary")
            monthly_display = room_cost_monthly_df.copy()
            for column in [
                "Water_Total", "Electricity_Total", "Linen_Laundry_Total", "SR_Total",
                "Total_Room_Cost", "Water_Per_Room", "Electricity_Per_Room",
                "Linen_Laundry_Per_Room", "SR_Per_Room", "Total_Cost_Per_Room",
            ]:
                monthly_display[column] = monthly_display[column].apply(_format_room_cost_value)
            st.dataframe(monthly_display, width="stretch", hide_index=True)

            st.markdown("#### Daily Breakdown")
            daily_display = room_cost_daily_df[[
                "Date", "Occupancy", "Room_Nights",
                "Guest_Supplies_Per_Room", "Cleaning_Supplies_Per_Room",
                "Paper_Supplies_Per_Room", "Print_Stationery_Per_Room", "Other_SR_Per_Room",
                "Water", "Water_Per_Room", "Electricity", "Electricity_Per_Room",
                "Linen_Laundry", "Linen_Laundry_Per_Room",
                "SR_Received_On_Date", "MTD_SR_Spend", "SR_Per_Room",
                "Total_Cost_Per_Room", "Total_Room_Cost",
            ]].copy()
            for column in daily_display.columns:
                if column == "Occupancy":
                    daily_display[column] = daily_display[column].apply(
                        lambda value: f"{value:.2f}%" if pd.notna(value) else ""
                    )
                elif column not in {"Date", "Room_Nights"}:
                    daily_display[column] = daily_display[column].apply(_format_room_cost_value)
            st.dataframe(daily_display, width="stretch", hide_index=True)
