"""
app.py - Streamlit Web Application for Hotel Santika Depok
Interactive Expense Separator & Spending Tracker.
"""

import os
import io
import numpy as np
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

import parser
import matcher
import exporter

# Configure Streamlit page
st.set_page_config(
    page_title="Hotel Spending Tracker & Separator",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS styling for modern executive appearance
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .badge-over {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 4px 8px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-under {
        background-color: #DCFCE7;
        color: #166534;
        padding: 4px 8px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        border-radius: 6px 6px 0 0;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Helper formatting
def format_idr(val):
    if pd.isna(val):
        return "Rp 0"
    return f"Rp {val:,.0f}"

def format_pct(val):
    if pd.isna(val):
        return "0.0%"
    return f"{val:+.1f}%"


# Base directories for standard data
BASE_REVIEW_DIR = os.environ.get("HOTEL_DATA_DIR", "/home/rzl/Documents/Business Review")
MONTH_FOLDERS = {
    "August 2026 (Agustus)": os.path.join(BASE_REVIEW_DIR, "8.AGUSTUS"),
    "July 2026 (Juli)": os.path.join(BASE_REVIEW_DIR, "7.JULY"),
}
has_local_folder = os.path.exists(BASE_REVIEW_DIR)

# Sidebar navigation & data selection
st.sidebar.image("https://img.icons8.com/color/96/hotel-star.png", width=64)
st.sidebar.title("Spending Separator")
st.sidebar.markdown("**Hotel Santika Depok**\n*Room Division & Housekeeping*")
st.sidebar.divider()

data_source = st.sidebar.radio(
    "Select Data Source",
    ["Select Month Folder", "Upload Custom Excel Files"],
    index=0 if has_local_folder else 1
)

reconciled_data = None
selected_month_name = "Agustus 2026"

if data_source == "Select Month Folder":
    month_choice = st.sidebar.selectbox("Choose Month", list(MONTH_FOLDERS.keys()))
    selected_month_name = month_choice.split()[0] + " 2026"
    month_dir = MONTH_FOLDERS[month_choice]

    if os.path.exists(month_dir):
        files = parser.find_month_files(month_dir)
        st.sidebar.markdown(f"**Loaded Files:**")
        st.sidebar.caption(f"📁 DTB: {'✅ Found' if files['dtb'] else '❌ Missing'}")
        st.sidebar.caption(f"📁 I/S MTD: {'✅ Found' if files['is_mtd'] else '❌ Missing'}")
        st.sidebar.caption(f"📁 Consumption: {'✅ Found' if files['consumption'] else '❌ Missing'}")

        if files['dtb']:
            with st.spinner("Processing & reconciling financial records..."):
                dtb_df = parser.parse_detail_trial_balance(files['dtb'])
                is_df = parser.parse_income_statement(files['is_mtd']) if files['is_mtd'] else pd.DataFrame()
                cons_df = parser.parse_consumption_report(files['consumption']) if files['consumption'] else pd.DataFrame()
                reconciled_data = matcher.reconcile_monthly_expenses(dtb_df, is_df, cons_df)
    else:
        st.sidebar.error(f"Directory not found: {month_dir}")

else:
    st.sidebar.markdown("**Upload Monthly Excel Files:**")
    dtb_file = st.sidebar.file_uploader("1. Detail Trial Balance (DTB)", type=["xlsx"])
    is_file = st.sidebar.file_uploader("2. Income Statement Dept (MTD)", type=["xlsx"])
    cons_file = st.sidebar.file_uploader("3. Consumption Report", type=["xlsx"])
    selected_month_name = st.sidebar.text_input("Month Label", "Custom Month")

    if dtb_file:
        with st.spinner("Parsing uploaded files..."):
            # Save temporary files to load via openpyxl
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
            reconciled_data = matcher.reconcile_monthly_expenses(dtb_df, is_df, cons_df)

            # Cleanup
            try:
                os.unlink(dtb_path)
                if is_path: os.unlink(is_path)
                if cons_path: os.unlink(cons_path)
            except Exception:
                pass


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
        use_container_width=True,
        type="primary"
    )

    # Main Page Header
    st.markdown('<div class="main-header">Hotel Expense Separator & Spending Tracker</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Hotel Santika Depok • Room Division & Housekeeping • Period: <b>{selected_month_name}</b></div>', unsafe_allow_html=True)

    # Top KPI Metrics row
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Spent", format_idr(metrics['total_spent']))
    with col2:
        st.metric("Total Budget", format_idr(metrics['total_budget']))
    with col3:
        var_color = "normal" if metrics['variance_idr'] <= 0 else "inverse"
        st.metric("Net Variance", format_idr(metrics['variance_idr']), delta=f"{metrics['variance_pct']:.1f}%", delta_color=var_color)
    with col4:
        st.metric("Reconciled Transactions", f"{metrics['transaction_count']:,} txns")
    with col5:
        budget_pct = (metrics['total_spent'] / metrics['total_budget'] * 100) if metrics['total_budget'] > 0 else 0
        st.metric("Budget Utilization", f"{budget_pct:.1f}%")

    st.markdown("<br>", unsafe_allow_html=True)

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
            # Donut chart
            fig_pie = px.pie(
                cat_df,
                values="Actual",
                names="Category",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Prism
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20), showlegend=False, height=360)
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_right:
            st.subheader("Budget vs. Actual (Top 8 Categories)")
            top_cats = cat_df.head(8).sort_values(by="Actual", ascending=True)
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                y=top_cats['Category'],
                x=top_cats['Budget'],
                name='Budget',
                orientation='h',
                marker=dict(color='#94A3B8')
            ))
            fig_bar.add_trace(go.Bar(
                y=top_cats['Category'],
                x=top_cats['Actual'],
                name='Actual',
                orientation='h',
                marker=dict(color='#2563EB')
            ))
            fig_bar.update_layout(
                barmode='group',
                margin=dict(t=20, b=20, l=20, r=20),
                height=360,
                xaxis_tickformat=',',
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        st.divider()

        st.subheader("Top 10 Largest Cost Drivers Across Hotel")
        top_10 = top_items.head(10).sort_values(by="Total_Amount", ascending=True)
        fig_cost = px.bar(
            top_10,
            x="Total_Amount",
            y="Item_Name",
            color="Category",
            orientation="h",
            labels={"Total_Amount": "Amount (IDR)", "Item_Name": "Expense Item / Vendor"},
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_cost.update_layout(height=400, margin=dict(t=20, b=20, l=20, r=20), xaxis_tickformat=',')
        st.plotly_chart(fig_cost, use_container_width=True)

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
                filtered_items = filtered_items[filtered_items['Item_Name'].str.lower().str.contains(item_search.lower(), na=False)]

        if not filtered_items.empty:
            # Summary KPIs for selected items
            total_items_spend = filtered_items['Total_Amount'].sum()
            total_items_qty = filtered_items['Total_Qty'].sum()
            total_orders_count = filtered_items['Order_Count'].sum()
            unique_items_cnt = len(filtered_items)

            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            with kpi1:
                st.metric("Total Spend", format_idr(total_items_spend))
            with kpi2:
                st.metric("Total Units / Quantities", f"{total_items_qty:,.0f} units")
            with kpi3:
                st.metric("Total Repeated Orders", f"{total_orders_count} deliveries")
            with kpi4:
                st.metric("Unique Items Tracked", f"{unique_items_cnt} items")

            st.markdown("<br>", unsafe_allow_html=True)

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
                use_container_width=True,
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

                    stat1, stat2, stat3, stat4 = st.columns(4)
                    with stat1:
                        st.metric(f"Total {selected_item_name}", f"{item_tot_qty:,.0f} {item_unit}")
                    with stat2:
                        st.metric("Total Month Spend", format_idr(item_tot_cost))
                    with stat3:
                        st.metric("Avg Unit Price", format_idr(item_avg_cost))
                    with stat4:
                        st.metric("Total Deliveries / Vouchers", f"{len(item_trxs)} times")

                    # Timeline Chart of Deliveries
                    fig_timeline = px.bar(
                        item_trxs,
                        x='Date',
                        y='Qty',
                        hover_data=['Amount', 'Voucher_Ref', 'Unit_Price'],
                        title=f"Delivery Timeline for {selected_item_name} ({item_unit})",
                        labels={'Qty': f"Quantity Delivered ({item_unit})", 'Date': "Order / Issuing Date"},
                        color_discrete_sequence=['#2563EB']
                    )
                    fig_timeline.update_layout(height=280, margin=dict(t=35, b=20, l=20, r=20))
                    st.plotly_chart(fig_timeline, use_container_width=True)

                    # Detailed Voucher Table
                    st.markdown(f"**Individual Order Vouchers for {selected_item_name}:**")
                    voucher_view = item_trxs[['Date', 'Voucher_Ref', 'Qty', 'Unit', 'Unit_Price', 'Amount', 'JRNL', 'Raw_Description']].copy()
                    voucher_view['Unit_Price'] = voucher_view['Unit_Price'].apply(format_idr)
                    voucher_view['Amount'] = voucher_view['Amount'].apply(format_idr)
                    st.dataframe(voucher_view, use_container_width=True)

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
            q = search_query.lower()
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
            use_container_width=True,
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
            use_container_width=True,
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
        st.markdown("Track expense shifts between **August 2026** and **July 2026**.")

        # Load July data if available
        july_dir = MONTH_FOLDERS.get("July 2026 (Juli)")
        aug_dir = MONTH_FOLDERS.get("August 2026 (Agustus)")

        if july_dir and os.path.exists(july_dir) and aug_dir and os.path.exists(aug_dir):
            with st.spinner("Calculating MoM trends..."):
                july_files = parser.find_month_files(july_dir)
                aug_files = parser.find_month_files(aug_dir)

                j_dtb = parser.parse_detail_trial_balance(july_files['dtb'])
                j_is = parser.parse_income_statement(july_files['is_mtd'])
                j_cons = parser.parse_consumption_report(july_files['consumption'])
                j_rec = matcher.reconcile_monthly_expenses(j_dtb, j_is, j_cons)

                a_cat = cat_df[['Category', 'Actual']].rename(columns={'Actual': 'August_Actual'})
                j_cat = j_rec['category_summary'][['Category', 'Actual']].rename(columns={'Actual': 'July_Actual'})

                mom_df = pd.merge(a_cat, j_cat, on="Category", how="outer").fillna(0.0)
                mom_df['MoM_Diff'] = mom_df['August_Actual'] - mom_df['July_Actual']
                mom_df['MoM_Pct'] = (mom_df['MoM_Diff'] / mom_df['July_Actual'] * 100.0).replace([np.inf, -np.inf], 0.0).fillna(0.0)
                mom_df = mom_df.sort_values(by="August_Actual", ascending=False)

            # MoM KPI overview
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("August 2026 Total", format_idr(metrics['total_spent']))
            with m2:
                july_total = j_rec['metrics']['total_spent']
                st.metric("July 2026 Total", format_idr(july_total))
            with m3:
                net_mom = metrics['total_spent'] - july_total
                net_mom_pct = (net_mom / july_total * 100.0) if july_total > 0 else 0
                st.metric("MoM Spending Shift", format_idr(net_mom), delta=f"{net_mom_pct:+.1f}%", delta_color="inverse")

            # MoM Comparison Bar Chart
            st.markdown("#### Top Categories Comparison (August vs July)")
            mom_top = mom_df.head(8)
            fig_mom = go.Figure()
            fig_mom.add_trace(go.Bar(
                x=mom_top['Category'],
                y=mom_top['July_Actual'],
                name='July 2026',
                marker=dict(color='#94A3B8')
            ))
            fig_mom.add_trace(go.Bar(
                x=mom_top['Category'],
                y=mom_top['August_Actual'],
                name='August 2026',
                marker=dict(color='#2563EB')
            ))
            fig_mom.update_layout(
                barmode='group',
                height=380,
                margin=dict(t=20, b=20, l=20, r=20),
                yaxis_tickformat=','
            )
            st.plotly_chart(fig_mom, use_container_width=True)

            # MoM Table
            mom_display = mom_df.copy()
            mom_display['August_Actual'] = mom_display['August_Actual'].apply(format_idr)
            mom_display['July_Actual'] = mom_display['July_Actual'].apply(format_idr)
            mom_display['MoM_Diff'] = mom_display['MoM_Diff'].apply(format_idr)
            mom_display['MoM_Pct'] = mom_display['MoM_Pct'].apply(format_pct)

            st.dataframe(
                mom_display,
                use_container_width=True,
                height=350,
                column_config={
                    "Category": st.column_config.TextColumn("Expense Category", width="large"),
                    "July_Actual": st.column_config.TextColumn("July 2026"),
                    "August_Actual": st.column_config.TextColumn("August 2026"),
                    "MoM_Diff": st.column_config.TextColumn("MoM Change (IDR)"),
                    "MoM_Pct": st.column_config.TextColumn("MoM Change (%)"),
                }
            )
        else:
            st.info("Previous month data not found for MoM comparison.")

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
