import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import requests
from io import BytesIO

# =====================================
# PAGE CONFIG
# =====================================
page = st.sidebar.radio(
    "📑 Select Report",
    ["Sales Contract", "Quotation", "Retention"]
)
if page == "Sales Contract":
    st.title("🚗 BMW Sales Contract Dashboard")
elif page == "Quotation":
    st.title("📋 BMW Quotation Dashboard")
else:
    st.title("💰 BMW Retention & Profitability Dashboard")

# =====================================
# CUSTOM CSS
# =====================================
st.markdown("""
<style>
.kpi-card {
    padding: 16px 10px;
    border-radius: 15px;
    text-align: center;
    color: white;
    margin-bottom: 12px;
    min-height: 110px;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    overflow: hidden;
}
.kpi-card h4 {
    margin: 0 0 6px 0;
    font-size: 0.85rem;
    font-weight: 500;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 100%;
}
.kpi-card h1 {
    margin: 0;
    font-size: 1.7rem;
    font-weight: 700;
    line-height: 1.1;
    white-space: nowrap;
}
.kpi-card h2 {
    margin: 0;
    font-size: 1.25rem;
    font-weight: 700;
    line-height: 1.15;
    white-space: nowrap;
}
.green  { background: linear-gradient(135deg,#16a34a,#15803d); }
.blue   { background: linear-gradient(135deg,#2563eb,#1d4ed8); }
.orange { background: linear-gradient(135deg,#ea580c,#c2410c); }
.purple { background: linear-gradient(135deg,#7c3aed,#6d28d9); }
.red    { background: linear-gradient(135deg,#dc2626,#b91c1c); }
.teal   { background: linear-gradient(135deg,#0d9488,#0f766e); }
.sky    { background: linear-gradient(135deg,#0284c7,#0369a1); }
.rose   { background: linear-gradient(135deg,#e11d48,#be123c); }
.indigo { background: linear-gradient(135deg,#4f46e5,#4338ca); }
.amber  { background: linear-gradient(135deg,#d97706,#b45309); }
.section-header {
    background: linear-gradient(90deg,#1e293b,#334155);
    color: white;
    padding: 12px 20px;
    border-radius: 10px;
    margin: 20px 0 10px 0;
    font-size: 1.1rem;
    font-weight: 600;
}
@media (max-width: 768px) {
    .kpi-card h1 { font-size: 1.3rem; }
    .kpi-card h2 { font-size: 1.05rem; }
    .kpi-card h4 { font-size: 0.75rem; }
}
</style>
""", unsafe_allow_html=True)

# =====================================
# ONEDRIVE DIRECT DOWNLOAD URLS
# =====================================
# FORMAT: your SharePoint share link + "?download=1" at the end
# Replace the Sales Contract and Quotation URLs once you have them
ONEDRIVE_URLS = {
    "Retention":      "https://infinitycr1-my.sharepoint.com/:x:/g/personal/pravinkumar_m_bmw-infinitycars_in/IQA5uS4P5kbFT4aI6nJSETv4AU6mCIA1K6rMgWGz_n3waVA?download=1",
    "Sales Contract": "https://infinitycr1-my.sharepoint.com/:x:/g/personal/pravinkumar_m_bmw-infinitycars_in/IQAUL6CF_KLYRrPoU9qAv9wuAbf9c3uJxKZ4F7-Vz7Lu7Kc?download=1",
    "Quotation":      "https://infinitycr1-my.sharepoint.com/:x:/g/personal/pravinkumar_m_bmw-infinitycars_in/IQBJgr2ET1LxTa5Uhx5EMx3CAUZDADqpCKAwZ9GI9R1MSCg?download=1",
}

# =====================================
# LOAD DATA FROM ONEDRIVE
# =====================================
@st.cache_data(ttl=300)  # Cache for 5 minutes — refresh button below clears it
def load_excel_from_onedrive(url: str) -> pd.DataFrame:
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return pd.read_excel(BytesIO(response.content), engine="openpyxl")

try:
    df = load_excel_from_onedrive(ONEDRIVE_URLS[page])

    if "Quotation Date" in df.columns:
        df["Quotation Date"] = pd.to_datetime(df["Quotation Date"], errors="coerce")

    # ── Retention-specific type coercion ──────────────────────────────
    if page == "Retention":
        if "Allot Dt" in df.columns:
            df["Allot Dt"] = pd.to_datetime(df["Allot Dt"], errors="coerce")
        if "Booking Month" in df.columns:
            df["Booking Month"] = pd.to_datetime(df["Booking Month"], errors="coerce")
        for col in ["Flexi Value", "Corporate Value", "Trade In/ Exchange Value"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        numeric_cols = [
            "Ex Showroom", "Whosale With Gst", "Basic Price", "Margin With GST",
            "Matrix Discount", "Insurance Discount", "SPL Discount Mgmt",
            "Retail Discount", "Ctc", "Ex Showroom Offered",
            "Profit With Gst", "Profit W/O Gst", "Vd Without Gst",
            "Retention W/O Vd", "Retention With Vd", "Vd Percentage",
            "My 2024", "Rm Kitty", "Corp Rm", "Exchange Bonus",
            "Retail Tgt Ach", "Inhouse Finance Support", "Psv Charges Without Gst",
            "Dsa Comission", "Additional Discount", "3rd Year Warranty",
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

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
    # RETENTION-SPECIFIC FILTERS
    # -------------------------------------
    if page == "Retention":
        if "Series" in df.columns:
            series_list = sorted(df["Series"].dropna().unique().tolist())
            selected_series = st.sidebar.multiselect("🔢 Series", options=series_list, default=[])
            if selected_series:
                filtered_df = filtered_df[filtered_df["Series"].isin(selected_series)]
        if "Sales Consultant" in df.columns:
            sc_list = sorted(df["Sales Consultant"].dropna().unique().tolist())
            selected_sc = st.sidebar.multiselect("👨‍💼 Sales Consultant", options=sc_list, default=[])
            if selected_sc:
                filtered_df = filtered_df[filtered_df["Sales Consultant"].isin(selected_sc)]
        if "Color" in df.columns:
            color_list = sorted(df["Color"].dropna().unique().tolist())
            selected_color = st.sidebar.multiselect("🎨 Color", options=color_list, default=[])
            if selected_color:
                filtered_df = filtered_df[filtered_df["Color"].isin(selected_color)]
        if "Interior" in df.columns:
            int_list = sorted(df["Interior"].dropna().unique().tolist())
            selected_int = st.sidebar.multiselect("🪑 Interior", options=int_list, default=[])
            if selected_int:
                filtered_df = filtered_df[filtered_df["Interior"].isin(selected_int)]
        if "Flexi" in df.columns:
            flexi_opts = ["All"] + sorted(df["Flexi"].dropna().unique().tolist())
            selected_flexi = st.sidebar.selectbox("💳 Flexi", options=flexi_opts)
            if selected_flexi != "All":
                filtered_df = filtered_df[filtered_df["Flexi"] == selected_flexi]
        if "Corporate" in df.columns:
            corp_opts = ["All"] + sorted(df["Corporate"].dropna().unique().tolist())
            selected_corp_scheme = st.sidebar.selectbox("🏢 Corporate Scheme", options=corp_opts)
            if selected_corp_scheme != "All":
                filtered_df = filtered_df[filtered_df["Corporate"] == selected_corp_scheme]
        if "Trade In/ Exchange" in df.columns:
            trade_opts = ["All"] + sorted(df["Trade In/ Exchange"].dropna().unique().tolist())
            selected_trade = st.sidebar.selectbox("🔄 Trade In/Exchange", options=trade_opts)
            if selected_trade != "All":
                filtered_df = filtered_df[filtered_df["Trade In/ Exchange"] == selected_trade]
        if "Retention With Vd" in df.columns:
            retention_filter = st.sidebar.selectbox(
                "📈 Retention Filter", options=["All Deals", "Positive Retention", "Negative Retention"]
            )
            if retention_filter == "Positive Retention":
                filtered_df = filtered_df[filtered_df["Retention With Vd"] > 0]
            elif retention_filter == "Negative Retention":
                filtered_df = filtered_df[filtered_df["Retention With Vd"] < 0]

    # -------------------------------------
    # DATE FILTER (Relative Periods) — Quotation / Sales Contract
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
    # DATE FILTER — Retention (Allot Dt)
    # -------------------------------------
    if page == "Retention" and "Allot Dt" in df.columns and not df["Allot Dt"].dropna().empty:
        st.sidebar.markdown("---")
        ret_date_period = st.sidebar.selectbox(
            "📅 Allotment Period",
            options=["All Time", "This Month", "Last 30 Days", "Custom Range"]
        )
        today = pd.Timestamp.now().normalize()
        if ret_date_period == "This Month":
            filtered_df = filtered_df[filtered_df["Allot Dt"] >= today.replace(day=1)]
        elif ret_date_period == "Last 30 Days":
            filtered_df = filtered_df[filtered_df["Allot Dt"] >= today - pd.Timedelta(days=30)]
        elif ret_date_period == "Custom Range":
            min_date = df["Allot Dt"].dropna().min().date()
            max_date = df["Allot Dt"].dropna().max().date()
            selected_dates = st.sidebar.date_input("Select Allotment Range", value=(min_date, max_date))
            if len(selected_dates) == 2:
                start_date, end_date = selected_dates
                filtered_df = filtered_df[
                    (filtered_df["Allot Dt"] >= pd.Timestamp(start_date)) &
                    (filtered_df["Allot Dt"] <= pd.Timestamp(end_date))
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
    if "Sales Consultant" in filtered_df.columns:
        st.sidebar.metric("Consultants", filtered_df["Sales Consultant"].nunique())
    st.sidebar.container(height=50, border=False)

    # =====================================
    # ACTION BUTTON — clears cache & reloads fresh from OneDrive
    # =====================================
    if st.button("🔄 Refresh CRM Data"):
        st.cache_data.clear()
        st.rerun()

    st.divider()

    if filtered_df.empty:
        st.warning("No records match the current filters.")
        st.stop()

    # =====================================
    # ₹ FORMATTER (used by Retention page)
    # =====================================
    def fmt_inr(val):
        if pd.isna(val):
            return "—"
        val = float(val)
        if abs(val) >= 1e7:
            return f"₹{val/1e7:.2f} Cr"
        elif abs(val) >= 1e5:
            return f"₹{val/1e5:.2f} L"
        else:
            return f"₹{val:,.0f}"

    # =====================================================================
    #                     RETENTION PAGE
    # =====================================================================
    if page == "Retention":
        # ─── Helper: safe sum ────────────────────────────────────────────
        def safe_sum(col):
            return filtered_df[col].sum() if col in filtered_df.columns else 0

        def safe_mean(col):
            return filtered_df[col].mean() if col in filtered_df.columns else 0

        # ─── SECTION 0 : EXECUTIVE SUMMARY KPIs ─────────────────────────
        st.markdown('<div class="section-header">📊 Executive Summary</div>', unsafe_allow_html=True)

        total_deals        = len(filtered_df)
        total_retention_vd = safe_sum("Retention With Vd")
        total_retention_wo = safe_sum("Retention W/O Vd")
        total_vd           = safe_sum("Vd Without Gst")
        total_ex_sr        = safe_sum("Ex Showroom")
        total_ex_offered   = safe_sum("Ex Showroom Offered")
        total_ctc          = safe_sum("Ctc")
        avg_vd_pct         = safe_mean("Vd Percentage")
        positive_deals     = (filtered_df["Retention With Vd"] > 0).sum() if "Retention With Vd" in filtered_df.columns else 0
        negative_deals     = (filtered_df["Retention With Vd"] <= 0).sum() if "Retention With Vd" in filtered_df.columns else 0
        win_rate           = (positive_deals / total_deals * 100) if total_deals > 0 else 0
        avg_retention_deal = total_retention_vd / total_deals if total_deals > 0 else 0

        r1c1, r1c2, r1c3, r1c4 = st.columns(4)
        r1c1.markdown(f'<div class="kpi-card blue"><h4>🚗 Total Deals</h4><h1>{total_deals}</h1></div>', unsafe_allow_html=True)
        r1c2.markdown(f'<div class="kpi-card green"><h4>💰 Total Retention (incl. VD)</h4><h2>{fmt_inr(total_retention_vd)}</h2></div>', unsafe_allow_html=True)
        r1c3.markdown(f'<div class="kpi-card sky"><h4>🔓 Retention (excl. VD)</h4><h2>{fmt_inr(total_retention_wo)}</h2></div>', unsafe_allow_html=True)
        r1c4.markdown(f'<div class="kpi-card teal"><h4>🎯 Volume Discount (VD)</h4><h2>{fmt_inr(total_vd)}</h2></div>', unsafe_allow_html=True)

        r2c1, r2c2, r2c3, r2c4 = st.columns(4)
        r2c1.markdown(f'<div class="kpi-card {"green" if avg_retention_deal >= 0 else "red"}"><h4>📈 Avg Retention / Deal</h4><h2>{fmt_inr(avg_retention_deal)}</h2></div>', unsafe_allow_html=True)
        r2c2.markdown(f'<div class="kpi-card purple"><h4>📐 Avg VD %</h4><h1>{avg_vd_pct:.2f}%</h1></div>', unsafe_allow_html=True)
        r2c3.markdown(f'<div class="kpi-card green"><h4>✅ Positive Retention Deals</h4><h1>{positive_deals}</h1></div>', unsafe_allow_html=True)
        r2c4.markdown(f'<div class="kpi-card red"><h4>❌ Negative Retention Deals</h4><h1>{negative_deals}</h1></div>', unsafe_allow_html=True)

        st.markdown(f"""
        <div style="background:#1e293b;border-radius:12px;padding:14px 20px;margin:10px 0 6px 0;color:white;">
          <span style="font-size:0.9rem;font-weight:600;">🏆 Positive Retention Win Rate &nbsp;—&nbsp;
          <span style="color:{'#4ade80' if win_rate>=50 else '#f87171'};font-size:1.2rem;">{win_rate:.1f}%</span>
          &nbsp;({positive_deals} of {total_deals} deals)</span>
          <div style="background:#334155;border-radius:6px;height:12px;margin-top:8px;">
            <div style="background:{'#16a34a' if win_rate>=50 else '#dc2626'};width:{min(win_rate,100):.1f}%;height:12px;border-radius:6px;"></div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.divider()

        # ─── SECTION 1 : DEAL FLOW & ALLOTMENT TIMELINE ─────────────────
        if "Allot Dt" in filtered_df.columns and not filtered_df["Allot Dt"].dropna().empty:
            st.markdown('<div class="section-header">📅 Deal Flow & Allotment Timeline</div>', unsafe_allow_html=True)
            tl = filtered_df.dropna(subset=["Allot Dt"]).copy()
            tl["Week"] = tl["Allot Dt"].dt.to_period("W").apply(lambda r: r.start_time)
            weekly = tl.groupby("Week").agg(
                Deals=("Vin Number", "count"),
                Retention=("Retention With Vd", "sum"),
            ).reset_index()
            tl["Date"] = tl["Allot Dt"].dt.date
            daily = tl.groupby("Date").agg(
                Deals=("Vin Number", "count"),
                Retention=("Retention With Vd", "sum"),
            ).reset_index()

            tf1, tf2 = st.columns(2)
            with tf1:
                fig_daily = make_subplots(specs=[[{"secondary_y": True}]])
                fig_daily.add_trace(go.Bar(x=daily["Date"], y=daily["Deals"],
                                           name="Deals", marker_color="#3b82f6", opacity=0.75), secondary_y=False)
                fig_daily.add_trace(go.Scatter(x=daily["Date"], y=daily["Retention"],
                                               name="Retention (₹)", mode="lines+markers",
                                               line=dict(color="#16a34a", width=2.5),
                                               marker=dict(size=7)), secondary_y=True)
                fig_daily.update_layout(title="📆 Daily Deal Volume vs Retention", height=360,
                                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                        legend=dict(orientation="h", y=-0.2))
                fig_daily.update_yaxes(title_text="# Deals", secondary_y=False)
                fig_daily.update_yaxes(title_text="₹ Retention", secondary_y=True)
                st.plotly_chart(fig_daily, use_container_width=True)
            with tf2:
                daily_sorted = daily.sort_values("Date").copy()
                daily_sorted["Cumulative Retention"] = daily_sorted["Retention"].cumsum()
                daily_sorted["Cumulative Deals"] = daily_sorted["Deals"].cumsum()
                fig_cum = make_subplots(specs=[[{"secondary_y": True}]])
                fig_cum.add_trace(go.Scatter(x=daily_sorted["Date"], y=daily_sorted["Cumulative Retention"],
                                             name="Cumulative Retention", fill="tozeroy", fillcolor="rgba(16,163,74,0.15)",
                                             line=dict(color="#16a34a", width=2.5)), secondary_y=False)
                fig_cum.add_trace(go.Scatter(x=daily_sorted["Date"], y=daily_sorted["Cumulative Deals"],
                                             name="Cumulative Deals", mode="lines+markers",
                                             line=dict(color="#2563eb", width=2, dash="dot"),
                                             marker=dict(size=6)), secondary_y=True)
                fig_cum.update_layout(title="📈 Cumulative Retention & Deal Momentum", height=360,
                                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                      legend=dict(orientation="h", y=-0.2))
                fig_cum.update_yaxes(title_text="₹ Cumulative Retention", secondary_y=False)
                fig_cum.update_yaxes(title_text="Cumulative Deals", secondary_y=True)
                st.plotly_chart(fig_cum, use_container_width=True)
            st.divider()

        # ─── SECTION 2 : SERIES & MODEL MIX ────────────────────────────
        st.markdown('<div class="section-header">🚘 Series & Model Performance</div>', unsafe_allow_html=True)
        if "Series" in filtered_df.columns:
            series_grp = filtered_df.groupby("Series").agg(
                Deals=("Vin Number", "count"),
                Retention=("Retention With Vd", "sum"),
                AvgRetention=("Retention With Vd", "mean"),
                VdPct=("Vd Percentage", "mean"),
            ).reset_index().sort_values("Retention", ascending=False)

            sm1, sm2 = st.columns(2)
            with sm1:
                fig_ser = make_subplots(specs=[[{"secondary_y": True}]])
                fig_ser.add_trace(go.Bar(x=series_grp["Series"], y=series_grp["Deals"],
                                         name="Deals", marker_color="#3b82f6",
                                         text=series_grp["Deals"], textposition="outside"), secondary_y=False)
                fig_ser.add_trace(go.Scatter(x=series_grp["Series"], y=series_grp["AvgRetention"],
                                              name="Avg Retention/Deal", mode="markers+lines",
                                              marker=dict(size=12, color="#f97316"),
                                              line=dict(color="#f97316", width=2)), secondary_y=True)
                fig_ser.update_layout(title="Deals & Avg Retention by Series", height=380,
                                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                fig_ser.update_yaxes(title_text="# Deals", secondary_y=False)
                fig_ser.update_yaxes(title_text="Avg Retention (₹)", secondary_y=True)
                st.plotly_chart(fig_ser, use_container_width=True)
            with sm2:
                sorted_ser = series_grp.sort_values("Retention")
                bar_colors = ["#16a34a" if v >= 0 else "#dc2626" for v in sorted_ser["Retention"]]
                fig_ser2 = go.Figure(go.Bar(
                    x=sorted_ser["Retention"], y=sorted_ser["Series"], orientation="h",
                    marker_color=bar_colors,
                    text=[fmt_inr(v) for v in sorted_ser["Retention"]], textposition="outside"
                ))
                fig_ser2.add_vline(x=0, line_dash="dash", line_color="gray")
                fig_ser2.update_layout(title="Total Retention (incl. VD) by Series", height=380,
                                       plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_ser2, use_container_width=True)

            if "Model" in filtered_df.columns:
                model_tree = filtered_df.groupby(["Series", "Model"]).agg(
                    Deals=("Vin Number", "count"),
                    Retention=("Retention With Vd", "sum"),
                ).reset_index()
                fig_tm = px.treemap(
                    model_tree, path=["Series", "Model"], values="Deals",
                    color="Retention", color_continuous_scale=["#dc2626", "#fbbf24", "#16a34a"],
                    title="📦 Model Mix — tile size = # Deals, colour = Total Retention"
                )
                fig_tm.update_layout(height=420)
                st.plotly_chart(fig_tm, use_container_width=True)
        st.divider()

        # ─── SECTION 3 : RETENTION WATERFALL (Deal Economics) ───────────
        st.markdown('<div class="section-header">💧 Deal Economics — How Retention is Built</div>', unsafe_allow_html=True)
        wf_components = {
            "Margin\n(with GST)":  safe_sum("Margin With GST"),
            "(-) Matrix\nDiscount": -abs(safe_sum("Matrix Discount")),
            "(-) Insurance\nDiscount": -abs(safe_sum("Insurance Discount")),
            "(-) SPL Disc\nMgmt": -abs(safe_sum("SPL Discount Mgmt")),
            "(-) Inhouse\nFinance": -abs(safe_sum("Inhouse Finance Support")),
            "(+) VD\n(Volume Disc)": safe_sum("Vd Without Gst"),
            "(+) Exchange\nBonus": safe_sum("Exchange Bonus"),
            "Net Retention\n(+VD)": safe_sum("Retention With Vd"),
        }
        wf_df = pd.DataFrame({"Component": list(wf_components.keys()), "Value": list(wf_components.values())})
        wf_colors = ["#16a34a" if v >= 0 else "#dc2626" for v in wf_df["Value"]]
        wf_colors[-1] = "#2563eb"

        we1, we2 = st.columns([3, 2])
        with we1:
            fig_wf = go.Figure(go.Bar(
                x=wf_df["Component"], y=wf_df["Value"],
                marker_color=wf_colors,
                text=[fmt_inr(v) for v in wf_df["Value"]], textposition="outside"
            ))
            fig_wf.add_hline(y=0, line_dash="dash", line_color="gray")
            fig_wf.update_layout(title="Aggregated Retention Build-up (₹)", height=420,
                                 plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_wf, use_container_width=True)
        with we2:
            pos_components = {k: v for k, v in {
                "Margin (GST)": safe_sum("Margin With GST"),
                "VD": safe_sum("Vd Without Gst"),
                "Exchange Bonus": safe_sum("Exchange Bonus"),
                "Corp RM": safe_sum("Corp Rm"),
            }.items() if v > 0}
            if pos_components:
                fig_donut = px.pie(
                    names=list(pos_components.keys()),
                    values=list(pos_components.values()),
                    hole=0.52,
                    title="📊 Positive Drivers Composition",
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                fig_donut.update_traces(textinfo="percent+label")
                fig_donut.update_layout(height=420)
                st.plotly_chart(fig_donut, use_container_width=True)
        st.divider()

        # ─── SECTION 4 : CONSULTANT PERFORMANCE ─────────────────────────
        if "Sales Consultant" in filtered_df.columns:
            st.markdown('<div class="section-header">👨‍💼 Sales Consultant Performance</div>', unsafe_allow_html=True)
            sc_grp = filtered_df.groupby("Sales Consultant").agg(
                Deals=("Vin Number", "count"),
                TotalRetention=("Retention With Vd", "sum"),
                AvgRetention=("Retention With Vd", "mean"),
                AvgVdPct=("Vd Percentage", "mean"),
                PositiveDeals=("Retention With Vd", lambda x: (x > 0).sum()),
            ).reset_index()
            sc_grp["NegativeDeals"] = sc_grp["Deals"] - sc_grp["PositiveDeals"]
            sc_grp["WinRate"] = (sc_grp["PositiveDeals"] / sc_grp["Deals"] * 100).round(1)
            sc_grp["RetentionPerDeal"] = sc_grp["TotalRetention"] / sc_grp["Deals"]

            cp1, cp2 = st.columns(2)
            with cp1:
                sc_sorted = sc_grp.sort_values("TotalRetention")
                bar_clrs = ["#16a34a" if v >= 0 else "#dc2626" for v in sc_sorted["TotalRetention"]]
                fig_sc_ret = go.Figure(go.Bar(
                    x=sc_sorted["TotalRetention"], y=sc_sorted["Sales Consultant"],
                    orientation="h", marker_color=bar_clrs,
                    text=[fmt_inr(v) for v in sc_sorted["TotalRetention"]], textposition="outside"
                ))
                fig_sc_ret.add_vline(x=0, line_dash="dash", line_color="gray")
                fig_sc_ret.update_layout(title="💰 Total Retention by Consultant", height=420,
                                         plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_sc_ret, use_container_width=True)
            with cp2:
                fig_sc_stack = go.Figure()
                fig_sc_stack.add_bar(name="✅ Positive", x=sc_grp["Sales Consultant"],
                                     y=sc_grp["PositiveDeals"], marker_color="#16a34a",
                                     text=sc_grp["PositiveDeals"], textposition="inside")
                fig_sc_stack.add_bar(name="❌ Negative", x=sc_grp["Sales Consultant"],
                                     y=sc_grp["NegativeDeals"], marker_color="#dc2626",
                                     text=sc_grp["NegativeDeals"], textposition="inside")
                fig_sc_stack.update_layout(barmode="stack",
                                           title="✅ Positive vs ❌ Negative Retention Deals",
                                           xaxis_tickangle=-35, height=420,
                                           plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_sc_stack, use_container_width=True)

            fig_quad = px.scatter(
                sc_grp, x="WinRate", y="RetentionPerDeal", size="Deals",
                color="TotalRetention", color_continuous_scale=["#dc2626", "#fbbf24", "#16a34a"],
                text="Sales Consultant",
                title="🎯 Consultant Quadrant — Win Rate % vs Avg Retention per Deal (bubble = # Deals)"
            )
            fig_quad.add_hline(y=sc_grp["RetentionPerDeal"].median(), line_dash="dash",
                               line_color="#94a3b8", annotation_text="Median Retention/Deal")
            fig_quad.add_vline(x=50, line_dash="dash", line_color="#94a3b8", annotation_text="50% Win Rate")
            fig_quad.update_traces(textposition="top center")
            fig_quad.update_layout(height=440, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_quad, use_container_width=True)

            st.subheader("🏆 Consultant Leaderboard")
            lb = sc_grp.sort_values("TotalRetention", ascending=False)[
                ["Sales Consultant", "Deals", "PositiveDeals", "NegativeDeals",
                 "WinRate", "TotalRetention", "RetentionPerDeal", "AvgVdPct"]
            ].copy()
            lb["TotalRetention"]    = lb["TotalRetention"].apply(fmt_inr)
            lb["RetentionPerDeal"]  = lb["RetentionPerDeal"].apply(fmt_inr)
            lb["AvgVdPct"]          = lb["AvgVdPct"].round(2).astype(str) + "%"
            lb["WinRate"]           = lb["WinRate"].astype(str) + "%"
            lb.columns = ["Consultant", "Deals", "✅ Positive", "❌ Negative",
                          "Win Rate", "Total Retention", "Avg Retention/Deal", "Avg VD%"]
            lb = lb.reset_index(drop=True)
            lb.index += 1
            st.dataframe(lb, use_container_width=True)
            st.divider()

        # ─── SECTION 5 : DISCOUNT ANATOMY ───────────────────────────────
        st.markdown('<div class="section-header">🔬 Discount Anatomy — Where the Money Goes</div>', unsafe_allow_html=True)
        disc_map = {
            "Matrix Discount": "Matrix Disc",
            "Insurance Discount": "Insurance Disc",
            "SPL Discount Mgmt": "SPL Disc",
            "Inhouse Finance Support": "Inhouse Finance",
            "Exchange Bonus": "Exchange Bonus",
            "Corp Rm": "Corp RM",
            "My 2024": "MY 2024",
            "Rm Kitty": "RM Kitty",
            "3rd Year Warranty": "3yr Warranty",
        }
        disc_vals = {label: filtered_df[col].abs().sum()
                     for col, label in disc_map.items() if col in filtered_df.columns and filtered_df[col].abs().sum() > 0}
        disc_df = pd.DataFrame({"Type": list(disc_vals.keys()), "Total": list(disc_vals.values())}).sort_values("Total", ascending=False)

        if not disc_df.empty:
            da1, da2 = st.columns(2)
            with da1:
                fig_pie_d = px.pie(disc_df, names="Type", values="Total", hole=0.45,
                                   title="🍩 Discount Mix — Share of Each Type",
                                   color_discrete_sequence=px.colors.qualitative.Bold)
                fig_pie_d.update_traces(textinfo="percent+label")
                fig_pie_d.update_layout(height=400)
                st.plotly_chart(fig_pie_d, use_container_width=True)
            with da2:
                fig_bar_d = px.bar(disc_df, x="Total", y="Type", orientation="h",
                                   color="Total", color_continuous_scale="Reds",
                                   text=[fmt_inr(v) for v in disc_df["Total"]],
                                   title="📊 Discount Quantum by Type (₹)")
                fig_bar_d.update_layout(height=400, coloraxis_showscale=False,
                                        yaxis=dict(categoryorder="total ascending"),
                                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_bar_d, use_container_width=True)

        if "Series" in filtered_df.columns:
            agg_disc = {col: "sum" for col in ["Matrix Discount", "Insurance Discount", "SPL Discount Mgmt", "Inhouse Finance Support"]
                        if col in filtered_df.columns}
            if agg_disc:
                disc_by_series = filtered_df.groupby("Series").agg(agg_disc).fillna(0).reset_index()
                color_map = {"Matrix Discount": "#2563eb", "Insurance Discount": "#16a34a",
                             "SPL Discount Mgmt": "#ea580c", "Inhouse Finance Support": "#7c3aed"}
                fig_disc_ser = go.Figure()
                for col in agg_disc:
                    fig_disc_ser.add_bar(name=col, x=disc_by_series["Series"],
                                         y=disc_by_series[col], marker_color=color_map.get(col, "#888"))
                fig_disc_ser.update_layout(barmode="stack",
                                            title="📦 Stacked Discount by Series — which model costs most in discounts",
                                            height=380, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_disc_ser, use_container_width=True)
        st.divider()

        # ─── SECTION 6 : VD% ANALYSIS ───────────────────────────────────
        if "Vd Percentage" in filtered_df.columns and filtered_df["Vd Percentage"].notna().any():
            st.markdown('<div class="section-header">📐 Volume Discount (VD%) Deep Dive</div>', unsafe_allow_html=True)
            vd1, vd2 = st.columns(2)
            with vd1:
                fig_vd_hist = px.histogram(filtered_df, x="Vd Percentage", nbins=12,
                                           color_discrete_sequence=["#0d9488"],
                                           title="Distribution of VD% Across All Deals")
                avg_vd = filtered_df["Vd Percentage"].mean()
                fig_vd_hist.add_vline(x=avg_vd, line_dash="dot", line_color="#f97316",
                                      annotation_text=f"Avg: {avg_vd:.2f}%")
                fig_vd_hist.update_layout(height=360, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_vd_hist, use_container_width=True)
            with vd2:
                if "Series" in filtered_df.columns:
                    fig_vd_box = go.Figure()
                    for s in filtered_df["Series"].dropna().unique():
                        vals = filtered_df[filtered_df["Series"] == s]["Vd Percentage"].dropna()
                        if not vals.empty:
                            fig_vd_box.add_trace(go.Box(y=vals, name=s, boxmean=True))
                    fig_vd_box.update_layout(title="VD% Range by Series (Box Plot)", height=360,
                                             plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_vd_box, use_container_width=True)

            if "Retention With Vd" in filtered_df.columns:
                vd_scatter = filtered_df[["Vd Percentage", "Retention With Vd", "Retention W/O Vd", "Model", "Series"]].dropna(subset=["Vd Percentage", "Retention With Vd"])
                fig_vd_scat = px.scatter(
                    vd_scatter, x="Vd Percentage", y="Retention With Vd",
                    color="Series", size_max=14,
                    hover_data=["Model", "Retention W/O Vd"],
                    title="🔍 VD% vs Retention (incl. VD) — Higher VD drives Retention up?"
                )
                fig_vd_scat.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Break-even")
                fig_vd_scat.update_layout(height=400, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_vd_scat, use_container_width=True)
            st.divider()

        # ─── SECTION 7 : SCHEME & DEAL TYPE ANALYSIS ────────────────────
        st.markdown('<div class="section-header">🎁 Scheme & Deal Type — Flexi / Corporate / Trade-In</div>', unsafe_allow_html=True)
        sc7a, sc7b, sc7c = st.columns(3)
        with sc7a:
            if "Flexi" in filtered_df.columns:
                fl = filtered_df["Flexi"].fillna("No Flexi").value_counts().reset_index()
                fl.columns = ["Flexi", "Count"]
                fig_fl = px.pie(fl, names="Flexi", values="Count", hole=0.5,
                                title="💳 Flexi Scheme Split",
                                color_discrete_sequence=["#2563eb", "#94a3b8", "#f97316"])
                fig_fl.update_traces(textinfo="percent+label")
                st.plotly_chart(fig_fl, use_container_width=True)
        with sc7b:
            if "Corporate" in filtered_df.columns:
                cp = filtered_df["Corporate"].fillna("No Scheme").value_counts().reset_index()
                cp.columns = ["Scheme", "Count"]
                fig_cp = px.pie(cp, names="Scheme", values="Count", hole=0.5,
                                title="🏢 Corporate Scheme Split",
                                color_discrete_sequence=px.colors.qualitative.Set2)
                fig_cp.update_traces(textinfo="percent+label")
                st.plotly_chart(fig_cp, use_container_width=True)
        with sc7c:
            if "Trade In/ Exchange" in filtered_df.columns:
                ti = filtered_df["Trade In/ Exchange"].fillna("None").value_counts().reset_index()
                ti.columns = ["Type", "Count"]
                fig_ti = px.pie(ti, names="Type", values="Count", hole=0.5,
                                title="🔄 Trade In / Exchange Split",
                                color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_ti.update_traces(textinfo="percent+label")
                st.plotly_chart(fig_ti, use_container_width=True)

        scheme_rows = []
        for col, label in [("Flexi", "Flexi"), ("Corporate", "Corporate"), ("Trade In/ Exchange", "Trade-In")]:
            if col in filtered_df.columns:
                grp = filtered_df.groupby(
                    filtered_df[col].fillna(f"No {label}")
                )["Retention With Vd"].agg(["mean", "count"]).reset_index()
                grp.columns = [label, "Avg Retention", "Deals"]
                grp["Scheme Type"] = label
                grp.rename(columns={label: "Scheme Value"}, inplace=True)
                scheme_rows.append(grp)
        if scheme_rows:
            scheme_impact = pd.concat(scheme_rows, ignore_index=True)
            fig_scheme = px.bar(
                scheme_impact, x="Scheme Value", y="Avg Retention",
                color="Scheme Type", barmode="group",
                text=[fmt_inr(v) for v in scheme_impact["Avg Retention"]],
                title="📊 Average Retention (incl. VD) by Scheme — which schemes are most profitable?",
                color_discrete_sequence=["#2563eb", "#7c3aed", "#0d9488"]
            )
            fig_scheme.add_hline(y=0, line_dash="dash", line_color="gray")
            fig_scheme.update_layout(height=400, xaxis_tickangle=-20,
                                     plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_scheme, use_container_width=True)
        st.divider()

        # ─── SECTION 8 : COLOR & SPECIFICATION TRENDS ───────────────────
        st.markdown('<div class="section-header">🎨 Color & Specification Trends</div>', unsafe_allow_html=True)
        col8a, col8b = st.columns(2)
        with col8a:
            if "Color" in filtered_df.columns:
                color_grp = filtered_df.groupby("Color").agg(
                    Deals=("Vin Number", "count"),
                    AvgRetention=("Retention With Vd", "mean")
                ).reset_index().sort_values("Deals", ascending=False).head(15)
                fig_col = px.bar(color_grp, x="Color", y="Deals", color="AvgRetention",
                                 color_continuous_scale=["#dc2626", "#fbbf24", "#16a34a"],
                                 title="🎨 Top Exterior Colors — colour = Avg Retention",
                                 text="Deals")
                fig_col.update_layout(xaxis_tickangle=-35, height=380,
                                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_col, use_container_width=True)
        with col8b:
            if "Interior" in filtered_df.columns:
                int_grp = filtered_df.groupby("Interior").agg(
                    Deals=("Vin Number", "count"),
                    AvgRetention=("Retention With Vd", "mean")
                ).reset_index().sort_values("Deals", ascending=False).head(15)
                fig_int = px.bar(int_grp, x="Interior", y="Deals", color="AvgRetention",
                                 color_continuous_scale=["#dc2626", "#fbbf24", "#16a34a"],
                                 title="🪑 Interior Choices — colour = Avg Retention",
                                 text="Deals")
                fig_int.update_layout(xaxis_tickangle=-35, height=380,
                                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_int, use_container_width=True)
        st.divider()

        # ─── SECTION 9 : MODEL-LEVEL PROFITABILITY MATRIX ───────────────
        if "Model" in filtered_df.columns:
            st.markdown('<div class="section-header">🧩 Model-Level Retention Matrix</div>', unsafe_allow_html=True)
            model_mat = filtered_df.groupby("Model").agg(
                Deals=("Vin Number", "count"),
                TotalRetention=("Retention With Vd", "sum"),
                AvgRetention=("Retention With Vd", "mean"),
                PositiveDeals=("Retention With Vd", lambda x: (x > 0).sum()),
            ).reset_index()
            model_mat["WinRate"] = (model_mat["PositiveDeals"] / model_mat["Deals"] * 100).round(1)
            if "Vd Percentage" in filtered_df.columns:
                vd_model = filtered_df.groupby("Model")["Vd Percentage"].mean().reset_index()
                vd_model.columns = ["Model", "AvgVd"]
                model_mat = model_mat.merge(vd_model, on="Model", how="left")

            mm1, mm2 = st.columns(2)
            with mm1:
                top12 = model_mat.sort_values("TotalRetention").tail(12)
                bar_clrs_mm = ["#16a34a" if v >= 0 else "#dc2626" for v in top12["TotalRetention"]]
                fig_mm = go.Figure(go.Bar(
                    x=top12["TotalRetention"], y=top12["Model"], orientation="h",
                    marker_color=bar_clrs_mm,
                    text=[fmt_inr(v) for v in top12["TotalRetention"]], textposition="outside"
                ))
                fig_mm.add_vline(x=0, line_dash="dash", line_color="gray")
                fig_mm.update_layout(title="🏅 Total Retention by Model (Top 12)", height=420,
                                     plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_mm, use_container_width=True)
            with mm2:
                fig_mm2 = px.scatter(
                    model_mat, x="Deals", y="AvgRetention",
                    size="Deals", color="TotalRetention",
                    color_continuous_scale=["#dc2626", "#fbbf24", "#16a34a"],
                    text="Model",
                    title="Volume vs Avg Retention per Model"
                )
                fig_mm2.add_hline(y=0, line_dash="dash", line_color="gray")
                fig_mm2.update_traces(textposition="top center")
                fig_mm2.update_layout(height=420, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_mm2, use_container_width=True)

            st.subheader("📋 Full Model Retention Table")
            disp_mm = model_mat.sort_values("TotalRetention", ascending=False).copy()
            disp_mm["TotalRetention"] = disp_mm["TotalRetention"].apply(fmt_inr)
            disp_mm["AvgRetention"]   = disp_mm["AvgRetention"].apply(fmt_inr)
            disp_mm["WinRate"]        = disp_mm["WinRate"].astype(str) + "%"
            rename_mm = {"Model": "Model", "Deals": "Deals", "PositiveDeals": "✅ Positive",
                         "WinRate": "Win Rate", "TotalRetention": "Total Retention", "AvgRetention": "Avg Retention/Deal"}
            if "AvgVd" in disp_mm.columns:
                disp_mm["AvgVd"] = disp_mm["AvgVd"].round(2).astype(str) + "%"
                rename_mm["AvgVd"] = "Avg VD%"
            disp_mm = disp_mm.rename(columns=rename_mm).reset_index(drop=True)
            disp_mm.index += 1
            st.dataframe(disp_mm, use_container_width=True)
            st.divider()

        # ─── SECTION 10 : DISCOUNT EFFICIENCY BY CONSULTANT ─────────────
        disc_eff_cols = [c for c in ["Matrix Discount", "Insurance Discount", "SPL Discount Mgmt", "Inhouse Finance Support"]
                         if c in filtered_df.columns]
        if disc_eff_cols and "Sales Consultant" in filtered_df.columns:
            st.markdown('<div class="section-header">⚖️ Discount Efficiency — Retention per ₹ Discounted</div>', unsafe_allow_html=True)
            eff = filtered_df.copy()
            eff["TotalDiscount"] = eff[disc_eff_cols].fillna(0).sum(axis=1)
            eff_grp = eff.groupby("Sales Consultant").agg(
                TotalDiscount=("TotalDiscount", "sum"),
                TotalRetention=("Retention With Vd", "sum"),
                Deals=("Vin Number", "count"),
            ).reset_index()
            eff_grp["DiscountPerDeal"] = eff_grp["TotalDiscount"] / eff_grp["Deals"]
            eff_grp["RetentionPerDiscountRupee"] = eff_grp.apply(
                lambda r: r["TotalRetention"] / r["TotalDiscount"] if r["TotalDiscount"] > 0 else 0, axis=1
            )

            ef1, ef2 = st.columns(2)
            with ef1:
                fig_ef1 = px.scatter(
                    eff_grp, x="DiscountPerDeal", y="TotalRetention",
                    size="Deals", color="TotalRetention",
                    color_continuous_scale=["#dc2626", "#fbbf24", "#16a34a"],
                    text="Sales Consultant",
                    title="Discount Given vs Retention Generated — by Consultant"
                )
                fig_ef1.update_traces(textposition="top center")
                fig_ef1.add_hline(y=0, line_dash="dash", line_color="gray")
                fig_ef1.update_layout(height=400, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_ef1, use_container_width=True)
            with ef2:
                eff_sorted = eff_grp.sort_values("RetentionPerDiscountRupee")
                clrs_ef = ["#16a34a" if v >= 0 else "#dc2626" for v in eff_sorted["RetentionPerDiscountRupee"]]
                fig_ef2 = go.Figure(go.Bar(
                    x=eff_sorted["RetentionPerDiscountRupee"], y=eff_sorted["Sales Consultant"],
                    orientation="h", marker_color=clrs_ef,
                    text=[f"₹{v:.2f}" for v in eff_sorted["RetentionPerDiscountRupee"]], textposition="outside"
                ))
                fig_ef2.add_vline(x=0, line_dash="dash", line_color="gray")
                fig_ef2.update_layout(title="💡 Retention per ₹1 of Discount Given",
                                      height=400, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_ef2, use_container_width=True)
            st.divider()

        # ─── SECTION 11 : CORRELATION HEATMAP ───────────────────────────
        st.markdown('<div class="section-header">🧠 What Drives Retention — Correlation Analysis</div>', unsafe_allow_html=True)
        driver_cols = [c for c in [
            "Ex Showroom Offered", "Matrix Discount", "Insurance Discount",
            "SPL Discount Mgmt", "Vd Percentage", "Margin With GST",
            "Ctc", "Retention W/O Vd", "Retention With Vd",
        ] if c in filtered_df.columns]
        corr_df = filtered_df[driver_cols].dropna(how="all")
        if len(corr_df) >= 3 and "Retention With Vd" in corr_df.columns:
            corr_mat = corr_df.corr()
            cd1, cd2 = st.columns(2)
            with cd1:
                fig_heat = go.Figure(data=go.Heatmap(
                    z=corr_mat.values, x=corr_mat.columns, y=corr_mat.columns,
                    colorscale=[[0, "#dc2626"], [0.5, "#f8fafc"], [1, "#16a34a"]],
                    zmid=0, text=corr_mat.round(2).values, texttemplate="%{text}",
                    colorbar=dict(title="r")
                ))
                fig_heat.update_layout(title="🔥 Correlation Heatmap — All Deal Variables", height=460)
                st.plotly_chart(fig_heat, use_container_width=True)
            with cd2:
                ret_corr = corr_mat["Retention With Vd"].drop("Retention With Vd").sort_values()
                clrs_corr = ["#16a34a" if v >= 0 else "#dc2626" for v in ret_corr.values]
                fig_drv = go.Figure(go.Bar(
                    x=ret_corr.values, y=ret_corr.index, orientation="h",
                    marker_color=clrs_corr,
                    text=[f"{v:.2f}" for v in ret_corr.values], textposition="outside"
                ))
                fig_drv.add_vline(x=0, line_dash="dash", line_color="gray")
                fig_drv.update_layout(title="🎯 Correlation with Retention (incl. VD)",
                                      height=460, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_drv, use_container_width=True)
            st.caption("💡 Correlation shows direction & strength of linear relationship only — not causation. "
                       "+1 = rises with Retention; −1 = falls as Retention rises.")
        st.divider()

        # ─── SECTION 12 : DEAL-LEVEL DETAIL TABLE ───────────────────────
        st.subheader("📋 Deal-Level Detail")
        table_cols = [
            "Vin Number", "Model", "Series", "Color", "Interior",
            "Sales Consultant", "Allot Dt",
            "Ex Showroom", "Ex Showroom Offered", "Ctc",
            "Retention With Vd", "Retention W/O Vd", "Vd Percentage",
            "Margin With GST", "Matrix Discount", "Insurance Discount",
            "Flexi", "Corporate", "Trade In/ Exchange",
        ]
        avail = [c for c in table_cols if c in filtered_df.columns]
        detail = filtered_df[avail].copy()
        if "Allot Dt" in detail.columns:
            detail["Allot Dt"] = pd.to_datetime(detail["Allot Dt"]).dt.strftime("%Y-%m-%d")
            detail = detail.sort_values("Allot Dt", ascending=False)
        st.dataframe(detail, use_container_width=True, hide_index=True, height=450)
        csv_data = detail.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Export Filtered Data as CSV", data=csv_data,
                           file_name="retention_filtered.csv", mime="text/csv")

    # =====================================================================
    #              SALES CONTRACT / QUOTATION
    # =====================================================================
    else:
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

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="kpi-card green"><h4>{main_kpi_title}</h4><h1>{total_records}</h1></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="kpi-card blue"><h4>Customers</h4><h1>{unique_customers}</h1></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="kpi-card orange"><h4>Sales Persons</h4><h1>{total_salespersons}</h1></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="kpi-card purple"><h4>Models</h4><h1>{total_models}</h1></div>', unsafe_allow_html=True)

        k1, k2, k3, k4 = st.columns(4)
        k1.info(f"🏆 Top Model\n\n{top_model}")
        k2.success(f"👨‍💼 Top Sales Person\n\n{top_salesperson}")
        k3.warning(f"🏦 Top Finance\n\n{top_finance}")
        if page == "Sales Contract":
            k4.error(f"📞 Top Source\n\n{top_source}")
        else:
            k4.error(f"📋 Top Status\n\n{top_status}")

        # =====================================================================
        # QUOTATION ANALYTICS
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
        # SALES CONTRACT ANALYTICS
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
        # DETAILED RECORDS TABLE
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

# =====================================================================
# ERROR HANDLING
# =====================================================================
except requests.exceptions.HTTPError as e:
    st.error(f"❌ Could not download file from OneDrive. HTTP Error: {e}")
    st.info("💡 Make sure your OneDrive link is set to 'Anyone with the link can view' and ends with `?download=1`")
except requests.exceptions.ConnectionError:
    st.error("❌ Network error — could not reach OneDrive. Check your internet connection.")
except requests.exceptions.Timeout:
    st.error("❌ Request timed out. OneDrive took too long to respond. Try refreshing.")
except Exception as e:
    st.error(f"❌ Unexpected error: {e}")
    st.exception(e)
