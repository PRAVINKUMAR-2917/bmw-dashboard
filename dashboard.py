import streamlit as st
import pandas as pd
import plotly.express as px

# =====================================
# PAGE CONFIG
# =====================================
page = st.sidebar.radio(
    "📑 Select Report",
    ["Sales Contract", "Quotation"]
)

if page == "Sales Contract":
    st.title("🚗 BMW Sales Contract Dashboard")
else:
    st.title("📋 BMW Quotation Dashboard")

# =====================================
# CUSTOM CSS
# =====================================
st.markdown("""
<style>
.kpi-card {
    padding: 20px;
    border-radius: 15px;
    text-align: center;
    color: white;
    margin-bottom: 10px;
}
.green { background: #16a34a; }
.blue { background: #2563eb; }
.orange { background: #ea580c; }
.purple { background: #7c3aed; }
.red { background: #dc2626; }
.teal { background: #0d9488; }
</style>
""", unsafe_allow_html=True)

# =====================================
# LOAD DATA
# =====================================
if page == "Sales Contract":
    EXCEL_FILE = "reports/sales_contract.xlsx"
else:
    EXCEL_FILE = "reports/quotation.xlsx"

try:
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")

    if "Quotation Date" in df.columns:
        df["Quotation Date"] = pd.to_datetime(df["Quotation Date"], errors="coerce")

    filtered_df = df.copy()

    # =====================================
    # SIDEBAR FILTERS
    # =====================================
    st.sidebar.title(f"📊 {page} Filters")

    if "Sales Person" in df.columns:
        sales_person_list = sorted(df["Sales Person"].dropna().unique().tolist())
        selected_sales_persons = st.sidebar.multiselect("👨‍💼 Sales Person", options=sales_person_list, default=[])
        if selected_sales_persons:
            filtered_df = filtered_df[filtered_df["Sales Person"].isin(selected_sales_persons)]

    if "Model" in df.columns:
        model_list = sorted(df["Model"].dropna().unique().tolist())
        selected_models = st.sidebar.multiselect("🚘 Model", options=model_list, default=[])
        if selected_models:
            filtered_df = filtered_df[filtered_df["Model"].isin(selected_models)]

    # -------------------------------------
    # QUOTATION-SPECIFIC FILTERS
    # -------------------------------------
    if page == "Quotation":
        if "Status" in df.columns:
            status_list = sorted(df["Status"].dropna().unique().tolist())
            selected_status = st.sidebar.multiselect("📌 Status", options=status_list, default=[])
            if selected_status:
                filtered_df = filtered_df[filtered_df["Status"].isin(selected_status)]

        if "Category" in df.columns:
            category_list = sorted(df["Category"].dropna().unique().tolist())
            selected_category = st.sidebar.multiselect("🏷️ Category", options=category_list, default=[])
            if selected_category:
                filtered_df = filtered_df[filtered_df["Category"].isin(selected_category)]

        if "Finance Type" in df.columns:
            finance_list = sorted(df["Finance Type"].dropna().unique().tolist())
            selected_finance = st.sidebar.multiselect("🏦 Finance Type", options=finance_list, default=[])
            if selected_finance:
                filtered_df = filtered_df[filtered_df["Finance Type"].isin(selected_finance)]

        if "Bank" in df.columns:
            bank_list = sorted(df["Bank"].dropna().unique().tolist())
            selected_bank = st.sidebar.multiselect("🏛️ Bank", options=bank_list, default=[])
            if selected_bank:
                filtered_df = filtered_df[filtered_df["Bank"].isin(selected_bank)]

        if "Insurance Company" in df.columns:
            insurance_list = sorted(df["Insurance Company"].dropna().unique().tolist())
            selected_insurance = st.sidebar.multiselect("🛡️ Insurance Company", options=insurance_list, default=[])
            if selected_insurance:
                filtered_df = filtered_df[filtered_df["Insurance Company"].isin(selected_insurance)]

        if "Corporate Company" in df.columns:
            corp_list = sorted(df["Corporate Company"].dropna().unique().tolist())
            selected_corp = st.sidebar.multiselect("🏢 Corporate Company", options=corp_list, default=[])
            if selected_corp:
                filtered_df = filtered_df[filtered_df["Corporate Company"].isin(selected_corp)]

        if "Approval Person" in df.columns:
            approval_list = sorted(df["Approval Person"].dropna().unique().tolist())
            selected_approval = st.sidebar.multiselect("✅ Approval Person", options=approval_list, default=[])
            if selected_approval:
                filtered_df = filtered_df[filtered_df["Approval Person"].isin(selected_approval)]

        if "Exterior" in df.columns:
            ext_list = sorted(df["Exterior"].dropna().unique().tolist())
            selected_ext = st.sidebar.multiselect("🎨 Exterior Color", options=ext_list, default=[])
            if selected_ext:
                filtered_df = filtered_df[filtered_df["Exterior"].isin(selected_ext)]

        if "Interior" in df.columns:
            int_list = sorted(df["Interior"].dropna().unique().tolist())
            selected_int = st.sidebar.multiselect("🪑 Interior", options=int_list, default=[])
            if selected_int:
                filtered_df = filtered_df[filtered_df["Interior"].isin(selected_int)]

        if "GST Number" in df.columns:
            gst_filter = st.sidebar.selectbox("🧾 GST Filter", options=["All", "With GST", "Without GST"])
            if gst_filter == "With GST":
                filtered_df = filtered_df[filtered_df["GST Number"].notna()]
            elif gst_filter == "Without GST":
                filtered_df = filtered_df[filtered_df["GST Number"].isna()]

    # -------------------------------------
    # SALES CONTRACT-SPECIFIC FILTERS
    # -------------------------------------
    if page == "Sales Contract":
        if "Series" in df.columns:
            series_list = sorted(df["Series"].dropna().unique().tolist())
            selected_series = st.sidebar.multiselect("🔢 Series", options=series_list, default=[])
            if selected_series:
                filtered_df = filtered_df[filtered_df["Series"].isin(selected_series)]

        if "Finance Type" in df.columns:
            finance_list = sorted(df["Finance Type"].dropna().unique().tolist())
            selected_finance = st.sidebar.multiselect("🏦 Finance Type", options=finance_list, default=[])
            if selected_finance:
                filtered_df = filtered_df[filtered_df["Finance Type"].isin(selected_finance)]

        if "Bank Name" in df.columns:
            bank_list = sorted(df["Bank Name"].dropna().unique().tolist())
            selected_bank = st.sidebar.multiselect("🏛️ Bank", options=bank_list, default=[])
            if selected_bank:
                filtered_df = filtered_df[filtered_df["Bank Name"].isin(selected_bank)]

        if "Insurance Company" in df.columns:
            insurance_list = sorted(df["Insurance Company"].dropna().unique().tolist())
            selected_insurance = st.sidebar.multiselect("🛡️ Insurance Company", options=insurance_list, default=[])
            if selected_insurance:
                filtered_df = filtered_df[filtered_df["Insurance Company"].isin(selected_insurance)]

        if "Corporate Company" in df.columns:
            corp_list = sorted(df["Corporate Company"].dropna().unique().tolist())
            selected_corp = st.sidebar.multiselect("🏢 Corporate Company", options=corp_list, default=[])
            if selected_corp:
                filtered_df = filtered_df[filtered_df["Corporate Company"].isin(selected_corp)]

        if "Source of Enquiry" in df.columns:
            source_list = sorted(df["Source of Enquiry"].dropna().unique().tolist())
            selected_source = st.sidebar.multiselect("📞 Source of Enquiry", options=source_list, default=[])
            if selected_source:
                filtered_df = filtered_df[filtered_df["Source of Enquiry"].isin(selected_source)]

        if "Exterior Color" in df.columns:
            ext_list = sorted(df["Exterior Color"].dropna().unique().tolist())
            selected_ext = st.sidebar.multiselect("🎨 Exterior Color", options=ext_list, default=[])
            if selected_ext:
                filtered_df = filtered_df[filtered_df["Exterior Color"].isin(selected_ext)]

        if "Interior Color" in df.columns:
            int_list = sorted(df["Interior Color"].dropna().unique().tolist())
            selected_int = st.sidebar.multiselect("🪑 Interior Color", options=int_list, default=[])
            if selected_int:
                filtered_df = filtered_df[filtered_df["Interior Color"].isin(selected_int)]

        if "GST No" in df.columns:
            gst_filter = st.sidebar.selectbox("🧾 GST Filter", options=["All", "With GST", "Without GST"])
            if gst_filter == "With GST":
                filtered_df = filtered_df[filtered_df["GST No"].notna()]
            elif gst_filter == "Without GST":
                filtered_df = filtered_df[filtered_df["GST No"].isna()]

    # -------------------------------------
    # DATE FILTER (Relative Periods)
    # -------------------------------------
    if "Quotation Date" in df.columns and not df["Quotation Date"].dropna().empty:
        st.sidebar.markdown("---")
        date_period = st.sidebar.selectbox(
            "📅 Date Period",
            options=["All Time", "Today", "This Week", "This Month", "Last 30 Days", "Custom Range"]
        )

        today = pd.Timestamp.now().normalize()

        if date_period == "Today":
            filtered_df = filtered_df[filtered_df["Quotation Date"] == today]
        elif date_period == "This Week":
            start_week = today - pd.Timedelta(days=today.weekday())
            filtered_df = filtered_df[(filtered_df["Quotation Date"] >= start_week) & (filtered_df["Quotation Date"] <= today)]
        elif date_period == "This Month":
            start_month = today.replace(day=1)
            filtered_df = filtered_df[(filtered_df["Quotation Date"] >= start_month) & (filtered_df["Quotation Date"] <= today)]
        elif date_period == "Last 30 Days":
            start_30 = today - pd.Timedelta(days=30)
            filtered_df = filtered_df[(filtered_df["Quotation Date"] >= start_30) & (filtered_df["Quotation Date"] <= today)]
        elif date_period == "Custom Range":
            min_date = df["Quotation Date"].dropna().min().date()
            max_date = df["Quotation Date"].dropna().max().date()
            selected_dates = st.sidebar.date_input("Select Custom Range", value=(min_date, max_date))
            if len(selected_dates) == 2:
                start_date, end_date = selected_dates
                filtered_df = filtered_df[
                    (filtered_df["Quotation Date"] >= pd.Timestamp(start_date)) &
                    (filtered_df["Quotation Date"] <= pd.Timestamp(end_date))
                ]

    # -------------------------------------
    # SIDEBAR STATS
    # -------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.metric("Filtered Records", len(filtered_df))

    if "Model" in filtered_df.columns:
        st.sidebar.metric("Models", filtered_df["Model"].nunique())

    if "Sales Person" in filtered_df.columns:
        st.sidebar.metric("Sales Persons", filtered_df["Sales Person"].nunique())

    st.sidebar.container(height=50, border=False)

    # =====================================
    # ACTION BUTTON
    # =====================================
    if st.button("🔄 Refresh CRM Data"):
        st.rerun()

    st.divider()

    # =====================================
    # KPI VALUES (COMMON)
    # =====================================
    total_records = len(filtered_df)

    unique_customers = filtered_df["Customer Name"].nunique() if "Customer Name" in filtered_df.columns else 0
    total_models = filtered_df["Model"].nunique() if "Model" in filtered_df.columns else 0
    total_salespersons = filtered_df["Sales Person"].nunique() if "Sales Person" in filtered_df.columns else 0

    top_model = (
        filtered_df["Model"].mode().iloc[0]
        if "Model" in filtered_df.columns and not filtered_df.empty and not filtered_df["Model"].mode().empty
        else "-"
    )

    top_salesperson = (
        filtered_df["Sales Person"].mode().iloc[0]
        if "Sales Person" in filtered_df.columns and not filtered_df.empty and not filtered_df["Sales Person"].mode().empty
        else "-"
    )

    finance_col = "Finance Type" if "Finance Type" in filtered_df.columns else None
    top_finance = (
        filtered_df[finance_col].mode().iloc[0]
        if finance_col and not filtered_df.empty and not filtered_df[finance_col].mode().empty
        else "-"
    )

    main_kpi_title = "Total Contracts" if page == "Sales Contract" else "Total Quotations"

    if "Source of Enquiry" in filtered_df.columns:
        top_source = (
            filtered_df["Source of Enquiry"].mode().iloc[0]
            if not filtered_df["Source of Enquiry"].mode().empty else "-"
        )
    else:
        top_source = "-"

    if "Status" in filtered_df.columns:
        top_status = (
            filtered_df["Status"].mode().iloc[0]
            if not filtered_df["Status"].mode().empty else "-"
        )
    else:
        top_status = "-"

    # =====================================
    # KPI ROW 1
    # =====================================
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f'<div class="kpi-card green"><h4>{main_kpi_title}</h4><h1>{total_records}</h1></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="kpi-card blue"><h4>Customers</h4><h1>{unique_customers}</h1></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="kpi-card orange"><h4>Sales Persons</h4><h1>{total_salespersons}</h1></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="kpi-card purple"><h4>Models</h4><h1>{total_models}</h1></div>', unsafe_allow_html=True)

    # =====================================
    # KPI ROW 2
    # =====================================
    k1, k2, k3, k4 = st.columns(4)

    k1.info(f"🏆 Top Model\n\n{top_model}")
    k2.success(f"👨‍💼 Top Sales Person\n\n{top_salesperson}")
    k3.warning(f"🏦 Top Finance\n\n{top_finance}")

    if page == "Sales Contract":
        k4.error(f"📞 Top Source\n\n{top_source}")
    else:
        k4.error(f"📋 Top Status\n\n{top_status}")

    # =====================================================================
    # QUOTATION-SPECIFIC EDA / ANALYTICS SECTION
    # =====================================================================
    if page == "Quotation":

        st.divider()
        st.header("🔎 Quotation Analytics & Insights")

        e1, e2, e3, e4 = st.columns(4)

        approved_count = (
            filtered_df["Status"].astype(str).str.lower().eq("approved").sum()
            if "Status" in filtered_df.columns else 0
        )
        pending_count = (
            filtered_df["Status"].astype(str).str.lower().eq("pending").sum()
            if "Status" in filtered_df.columns else 0
        )
        rejected_count = (
            filtered_df["Status"].astype(str).str.lower().eq("rejected").sum()
            if "Status" in filtered_df.columns else 0
        )
        conversion_rate = round((approved_count / total_records) * 100, 1) if total_records > 0 else 0

        e1.markdown(f'<div class="kpi-card green"><h4>✅ Approved</h4><h1>{approved_count}</h1></div>', unsafe_allow_html=True)
        e2.markdown(f'<div class="kpi-card orange"><h4>⏳ Pending</h4><h1>{pending_count}</h1></div>', unsafe_allow_html=True)
        e3.markdown(f'<div class="kpi-card red"><h4>❌ Rejected</h4><h1>{rejected_count}</h1></div>', unsafe_allow_html=True)
        e4.markdown(f'<div class="kpi-card teal"><h4>📈 Conversion Rate</h4><h1>{conversion_rate}%</h1></div>', unsafe_allow_html=True)

        st.markdown("")

        r1c1, r1c2 = st.columns(2)
        with r1c1:
            if "Status" in filtered_df.columns and not filtered_df.empty:
                st.subheader("📌 Quotation Status Breakdown")
                status_counts = filtered_df["Status"].fillna("Unknown").value_counts().reset_index()
                status_counts.columns = ["Status", "Count"]
                fig = px.pie(status_counts, names="Status", values="Count", hole=0.45)
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

        with r1c2:
            if "Category" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🏷️ Category Distribution")
                cat_counts = filtered_df["Category"].fillna("Unknown").value_counts().reset_index()
                cat_counts.columns = ["Category", "Count"]
                fig = px.bar(cat_counts, x="Category", y="Count", color="Category", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            if "Model" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🚘 Top Models Quoted")
                model_counts = filtered_df["Model"].fillna("Unknown").value_counts().head(10).reset_index()
                model_counts.columns = ["Model", "Count"]
                fig = px.bar(model_counts, x="Count", y="Model", orientation="h", color="Count", color_continuous_scale="Blues", text="Count")
                fig.update_layout(yaxis=dict(categoryorder="total ascending"), coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)

        with r2c2:
            if "Finance Type" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🏦 Finance Type Split")
                finance_counts = filtered_df["Finance Type"].fillna("Unknown").value_counts().reset_index()
                finance_counts.columns = ["Finance Type", "Count"]
                fig = px.pie(finance_counts, names="Finance Type", values="Count", hole=0.45,
                              color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

        r3c1, r3c2 = st.columns(2)
        with r3c1:
            if "Sales Person" in filtered_df.columns and "Status" in filtered_df.columns and not filtered_df.empty:
                st.subheader("👨‍💼 Sales Person vs Status")
                pivot = filtered_df.groupby(["Sales Person", "Status"]).size().reset_index(name="Count")
                fig = px.bar(pivot, x="Sales Person", y="Count", color="Status", barmode="stack")
                st.plotly_chart(fig, use_container_width=True)

        with r3c2:
            if "Approval Person" in filtered_df.columns and not filtered_df.empty:
                st.subheader("✅ Approvals by Person")
                appr_counts = filtered_df["Approval Person"].fillna("Unassigned").value_counts().reset_index()
                appr_counts.columns = ["Approval Person", "Count"]
                fig = px.bar(appr_counts, x="Approval Person", y="Count", color="Approval Person", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        r4c1, r4c2 = st.columns(2)
        with r4c1:
            if "Bank" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🏛️ Bank-wise Quotations")
                bank_counts = filtered_df["Bank"].fillna("Unknown").value_counts().reset_index()
                bank_counts.columns = ["Bank", "Count"]
                fig = px.bar(bank_counts, x="Bank", y="Count", color="Bank", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        with r4c2:
            if "Insurance Company" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🛡️ Insurance Company-wise Quotations")
                ins_counts = filtered_df["Insurance Company"].fillna("Unknown").value_counts().reset_index()
                ins_counts.columns = ["Insurance Company", "Count"]
                fig = px.bar(ins_counts, x="Insurance Company", y="Count", color="Insurance Company", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        r5c1, r5c2 = st.columns(2)
        with r5c1:
            if "Exterior" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🎨 Popular Exterior Colors")
                ext_counts = filtered_df["Exterior"].fillna("Unknown").value_counts().head(10).reset_index()
                ext_counts.columns = ["Exterior", "Count"]
                fig = px.bar(ext_counts, x="Exterior", y="Count", color="Exterior", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        with r5c2:
            if "Interior" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🪑 Popular Interior Choices")
                int_counts = filtered_df["Interior"].fillna("Unknown").value_counts().head(10).reset_index()
                int_counts.columns = ["Interior", "Count"]
                fig = px.bar(int_counts, x="Interior", y="Count", color="Interior", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        if "Quotation Date" in filtered_df.columns and not filtered_df["Quotation Date"].dropna().empty:
            st.subheader("📈 Quotation Trend Over Time")
            trend_df = filtered_df.dropna(subset=["Quotation Date"]).copy()
            trend_df["Date"] = trend_df["Quotation Date"].dt.date
            trend_counts = trend_df.groupby("Date").size().reset_index(name="Quotations")
            fig = px.line(trend_counts, x="Date", y="Quotations", markers=True)
            st.plotly_chart(fig, use_container_width=True)

        r7c1, r7c2 = st.columns(2)
        with r7c1:
            if "GST Number" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🧾 GST vs Non-GST Customers")
                gst_split = filtered_df["GST Number"].notna().map({True: "With GST", False: "Without GST"}).value_counts().reset_index()
                gst_split.columns = ["GST Status", "Count"]
                fig = px.pie(gst_split, names="GST Status", values="Count", hole=0.45,
                              color_discrete_sequence=["#16a34a", "#dc2626"])
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

        with r7c2:
            if "Corporate Company" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🏢 Corporate vs Individual Customers")
                corp_split = filtered_df["Corporate Company"].notna().map({True: "Corporate", False: "Individual"}).value_counts().reset_index()
                corp_split.columns = ["Customer Type", "Count"]
                fig = px.pie(corp_split, names="Customer Type", values="Count", hole=0.45,
                              color_discrete_sequence=["#7c3aed", "#2563eb"])
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

    # =====================================================================
    # SALES CONTRACT-SPECIFIC EDA / ANALYTICS SECTION
    # =====================================================================
    else:
        st.divider()
        st.header("🔎 Sales Contract Analytics & Insights")

        s1, s2, s3, s4 = st.columns(4)

        finance_done = (
            filtered_df["Finance Type"].notna().sum()
            if "Finance Type" in filtered_df.columns else 0
        )
        cash_deals = (
            (filtered_df["Finance Type"].fillna("Cash").astype(str).str.lower() == "cash").sum()
            if "Finance Type" in filtered_df.columns else 0
        )
        corporate_deals = (
            filtered_df["Corporate Company"].notna().sum()
            if "Corporate Company" in filtered_df.columns else 0
        )
        with_insurance = (
            filtered_df["Insurance Company"].notna().sum()
            if "Insurance Company" in filtered_df.columns else 0
        )

        s1.markdown(f'<div class="kpi-card teal"><h4>💳 Financed Deals</h4><h1>{finance_done}</h1></div>', unsafe_allow_html=True)
        s2.markdown(f'<div class="kpi-card green"><h4>💵 Cash Deals</h4><h1>{cash_deals}</h1></div>', unsafe_allow_html=True)
        s3.markdown(f'<div class="kpi-card purple"><h4>🏢 Corporate Deals</h4><h1>{corporate_deals}</h1></div>', unsafe_allow_html=True)
        s4.markdown(f'<div class="kpi-card orange"><h4>🛡️ Insured Deals</h4><h1>{with_insurance}</h1></div>', unsafe_allow_html=True)

        st.markdown("")

        col_chart1, col_chart2 = st.columns(2)

        with col_chart1:
            if "Model" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🚘 Top Models Sold")
                model_counts = filtered_df["Model"].fillna("Unknown").value_counts().head(10).reset_index()
                model_counts.columns = ["Model", "Count"]
                fig = px.bar(model_counts, x="Count", y="Model", orientation="h", color="Count", color_continuous_scale="Blues", text="Count")
                fig.update_layout(yaxis=dict(categoryorder="total ascending"), coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)

            if "Finance Type" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🏦 Finance Type Distribution")
                finance_counts = filtered_df["Finance Type"].fillna("Cash/Unknown").value_counts().reset_index()
                finance_counts.columns = ["Finance Type", "Count"]
                fig = px.pie(finance_counts, names="Finance Type", values="Count", hole=0.45,
                              color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

            if "Series" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🔢 Series-wise Sales")
                series_counts = filtered_df["Series"].fillna("Unknown").value_counts().reset_index()
                series_counts.columns = ["Series", "Count"]
                fig = px.bar(series_counts, x="Series", y="Count", color="Series", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        with col_chart2:
            if "Sales Person" in filtered_df.columns and not filtered_df.empty:
                st.subheader("👨‍💼 Top Sales Persons")
                sales_counts = filtered_df["Sales Person"].fillna("Unknown").value_counts().head(10).reset_index()
                sales_counts.columns = ["Sales Person", "Count"]
                fig = px.bar(sales_counts, x="Sales Person", y="Count", color="Sales Person", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

            if "Source of Enquiry" in filtered_df.columns and not filtered_df.empty:
                st.subheader("📞 Source of Enquiry")
                source_counts = filtered_df["Source of Enquiry"].fillna("Unknown").value_counts().reset_index()
                source_counts.columns = ["Source", "Count"]
                fig = px.pie(source_counts, names="Source", values="Count", hole=0.45)
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

            if "Bank Name" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🏛️ Bank-wise Financing")
                bank_counts = filtered_df["Bank Name"].dropna().value_counts().reset_index()
                bank_counts.columns = ["Bank", "Count"]
                if not bank_counts.empty:
                    fig = px.bar(bank_counts, x="Bank", y="Count", color="Bank", text="Count")
                    fig.update_layout(showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)

        col_chart3, col_chart4 = st.columns(2)

        with col_chart3:
            if "Exterior Color" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🎨 Popular Exterior Colors")
                ext_counts = filtered_df["Exterior Color"].fillna("Unknown").value_counts().head(10).reset_index()
                ext_counts.columns = ["Exterior Color", "Count"]
                fig = px.bar(ext_counts, x="Exterior Color", y="Count", color="Exterior Color", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        with col_chart4:
            if "Interior Color" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🪑 Popular Interior Colors")
                int_counts = filtered_df["Interior Color"].fillna("Unknown").value_counts().head(10).reset_index()
                int_counts.columns = ["Interior Color", "Count"]
                fig = px.bar(int_counts, x="Interior Color", y="Count", color="Interior Color", text="Count")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        if "Quotation Date" in filtered_df.columns and not filtered_df["Quotation Date"].dropna().empty:
            st.subheader("📈 Sales Trend Over Time")
            trend_df = filtered_df.dropna(subset=["Quotation Date"]).copy()
            trend_df["Date"] = trend_df["Quotation Date"].dt.date
            trend_counts = trend_df.groupby("Date").size().reset_index(name="Sales")
            fig = px.line(trend_counts, x="Date", y="Sales", markers=True)
            st.plotly_chart(fig, use_container_width=True)

        col_chart5, col_chart6 = st.columns(2)

        with col_chart5:
            if "GST No" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🧾 GST vs Non-GST Customers")
                gst_split = filtered_df["GST No"].notna().map({True: "With GST", False: "Without GST"}).value_counts().reset_index()
                gst_split.columns = ["GST Status", "Count"]
                fig = px.pie(gst_split, names="GST Status", values="Count", hole=0.45,
                              color_discrete_sequence=["#16a34a", "#dc2626"])
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

        with col_chart6:
            if "Corporate Company" in filtered_df.columns and not filtered_df.empty:
                st.subheader("🏢 Corporate vs Individual Customers")
                corp_split = filtered_df["Corporate Company"].notna().map({True: "Corporate", False: "Individual"}).value_counts().reset_index()
                corp_split.columns = ["Customer Type", "Count"]
                fig = px.pie(corp_split, names="Customer Type", values="Count", hole=0.45,
                              color_discrete_sequence=["#7c3aed", "#2563eb"])
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(fig, use_container_width=True)

    # =====================================================================
    # DETAILED RECORDS TABLE (NO SENSITIVE CUSTOMER DATA)
    # =====================================================================
    st.divider()
    st.subheader("📋 Detailed Records")

    if page == "Quotation":
        safe_columns = [
            "Model", "Category", "Status", "Sales Person", "Finance Type",
            "Bank", "Insurance Company", "Interior", "Exterior", "Quotation Date"
        ]
    else:
        safe_columns = [
            "Series", "Model", "Interior Color", "Exterior Color",
            "Finance Type", "Bank Name", "Insurance Company",
            "Source of Enquiry", "Quotation Date"
        ]

    available_columns = [col for col in safe_columns if col in filtered_df.columns]
    summary_df = filtered_df[available_columns].copy()

    if "Quotation Date" in summary_df.columns:
        summary_df = summary_df.sort_values(by="Quotation Date", ascending=False)

    st.dataframe(summary_df, use_container_width=True, hide_index=True)

except FileNotFoundError:
    st.error(f"❌ File not found: {EXCEL_FILE}")


