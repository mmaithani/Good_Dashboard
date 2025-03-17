import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ----- PAGE CONFIG -----
st.set_page_config(
    page_title="SuperStore KPI Dashboard",
    layout="wide",
    page_icon=":bar_chart:"
)

# ----- LOAD EXTERNAL FONTAWESOME FOR ICONS -----
st.markdown(
    """
    <link rel="stylesheet" 
          href="https://use.fontawesome.com/releases/v5.15.4/css/all.css" 
          integrity="sha384-DyZ88mC6kzNeFjsV12o4F2X0p7mUp72mmfj/wzLA16pJo1sw8a4z9ShS4rA6m1wR" 
          crossorigin="anonymous">
    """,
    unsafe_allow_html=True
)

# ----- CUSTOM CSS -----
st.markdown(
    """
    <style>
    /* Hide Streamlit default elements for a cleaner look */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Custom scrollbar styling */
    ::-webkit-scrollbar {
        width: 8px;
    }
    ::-webkit-scrollbar-track {
        background: #ccc;
    }
    ::-webkit-scrollbar-thumb {
        background: #888;
    }

    /* KPI box styling */
    .kpi-box {
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid #444;
        border-radius: 8px;
        padding: 16px;
        margin: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        transition: transform 0.3s ease-in-out;
    }
    .kpi-box:hover {
        transform: scale(1.05);
    }
    .kpi-title {
        font-weight: 600;
        color: #FFFFFF;
        font-size: 16px;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-weight: 700;
        font-size: 24px;
        color: #1E90FF;
    }
    .icon {
        font-size: 24px;
        margin-right: 8px;
        color: #1E90FF;
    }
    /* Tooltip styling */
    .tooltip-text {
        visibility: hidden;
        width: 220px;
        background-color: #333;
        color: #fff;
        text-align: center;
        border-radius: 6px;
        padding: 8px;
        position: absolute;
        z-index: 1;
        bottom: 110%;
        left: 50%;
        margin-left: -110px;
        opacity: 0;
        transition: opacity 0.3s;
        font-size: 14px;
    }
    .tooltip-text::after {
        content: "";
        position: absolute;
        top: 100%;
        left: 50%;
        margin-left: -5px;
        border-width: 5px;
        border-style: solid;
        border-color: #333 transparent transparent transparent;
    }
    .kpi-box:hover .tooltip-text {
        visibility: visible;
        opacity: 1;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ----- LOAD DATA WITH CACHING -----
@st.cache_data
def load_data():
    df = pd.read_excel("Sample - Superstore-1.xlsx",engine="openpyxl")
    # Convert "Order Date" to datetime if needed
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ----- PAGE TITLE -----
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>SuperStore KPI Dashboard</h1>", unsafe_allow_html=True)

# ----- SIDEBAR FILTERS & SETTINGS -----
st.sidebar.title("Filters & Settings")

# Dashboard Theme Option
theme_choice = st.sidebar.radio("Dashboard Theme", ["Dark", "Light"], index=0)
plotly_template = "plotly_dark" if theme_choice == "Dark" else "plotly_white"

# Aggregation Level Option
agg_level = st.sidebar.radio("Aggregation Level", ["Daily", "Weekly", "Monthly"], index=2)

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_regions = st.sidebar.multiselect("Region(s)", options=all_regions, default=all_regions)
df_filtered = df_original[df_original["Region"].isin(selected_regions)] if selected_regions else df_original

# State Filter
all_states = sorted(df_filtered["State"].dropna().unique())
selected_states = st.sidebar.multiselect("State(s)", options=all_states, default=all_states)
df_filtered = df_filtered[df_filtered["State"].isin(selected_states)] if selected_states else df_filtered

# Category Filter
all_categories = sorted(df_filtered["Category"].dropna().unique())
selected_categories = st.sidebar.multiselect("Category(ies)", options=all_categories, default=all_categories)
df_filtered = df_filtered[df_filtered["Category"].isin(selected_categories)] if selected_categories else df_filtered

# Sub-Category Filter
all_subcategories = sorted(df_filtered["Sub-Category"].dropna().unique())
selected_subcategories = st.sidebar.multiselect("Sub-Category(ies)", options=all_subcategories, default=all_subcategories)
df_filtered = df_filtered[df_filtered["Sub-Category"].isin(selected_subcategories)] if selected_subcategories else df_filtered

# Date Range Filter
if not df_filtered.empty:
    min_date = df_filtered["Order Date"].min()
    max_date = df_filtered["Order Date"].max()
else:
    min_date = df_original["Order Date"].min()
    max_date = df_original["Order Date"].max()

st.sidebar.subheader("Date Range")
from_date = st.sidebar.date_input("From", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("To", value=max_date, min_value=min_date, max_value=max_date)

if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

df = df_filtered[
    (df_filtered["Order Date"] >= pd.to_datetime(from_date)) &
    (df_filtered["Order Date"] <= pd.to_datetime(to_date))
].copy()

# ----- KPI CALCULATIONS -----
if df.empty:
    total_sales = total_quantity = total_profit = margin_rate = avg_discount = 0
else:
    total_sales = df["Sales"].sum()
    total_quantity = df["Quantity"].sum()
    total_profit = df["Profit"].sum()
    margin_rate = total_profit / total_sales if total_sales else 0
    avg_discount = df["Discount"].mean() if "Discount" in df.columns else 0

# Determine color for margin rate KPI (red if below 15% target)
margin_color = "#FF4B4B" if margin_rate < 0.15 else "#1E90FF"

# ----- KPI DISPLAY -----
kpi_cols = st.columns(5)
kpi_data = [
    ("Sales", f"${total_sales:,.2f}", "Total revenue generated."),
    ("Quantity Sold", f"{total_quantity:,.0f}", "Total units sold."),
    ("Profit", f"${total_profit:,.2f}", "Net profit after costs."),
    ("Margin Rate", f"{(margin_rate * 100):.2f}%", "Profit margin percentage.", margin_color),
    ("Avg Discount", f"{(avg_discount * 100):.2f}%", "Average discount applied.")
]

for i, item in enumerate(kpi_data):
    title, value, tooltip = item[0], item[1], item[2]
    extra_style = f"style='color:{item[3]}'" if len(item) > 3 else ""
    kpi_cols[i].markdown(
        f"""
        <div class='kpi-box' title='{tooltip}'>
            <div class='kpi-title'><i class="icon fa fa-{ 'dollar-sign' if title=='Sales' else 
                                                       'boxes' if title=='Quantity Sold' else 
                                                       'money-bill' if title=='Profit' else 
                                                       'percent' if title=='Margin Rate' else 
                                                       'tag' }"></i>{title}</div>
            <div class='kpi-value' {extra_style}>{value}</div>
            <span class="tooltip-text">{tooltip}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

# ----- ADD A GAUGE CHART FOR MARGIN RATE -----
st.markdown("---")
st.markdown("### Margin Rate Gauge")
fig_gauge = go.Figure(go.Indicator(
    mode="gauge+number+delta",
    value=margin_rate * 100,
    delta={'reference': 15, 'increasing': {'color': "red"}},
    gauge={
        'axis': {'range': [None, 100], 'tickprefix': "%"},
        'bar': {'color': margin_color},
        'steps': [
            {'range': [0, 15], 'color': "rgba(255,0,0,0.2)"},
            {'range': [15, 100], 'color': "rgba(0,255,0,0.2)"}
        ],
        'threshold': {
            'line': {'color': "orange", 'width': 4},
            'thickness': 0.75,
            'value': 15}
    },
    title={"text": "Margin Rate"}
))
st.plotly_chart(fig_gauge, use_container_width=True)

st.markdown("---")

# ----- CHARTS -----
if df.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    # KPI for chart selection
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate", "Avg Discount"]
    selected_kpi = st.radio("Select KPI for Time Series & Product Analysis:", options=kpi_options, horizontal=True)

    # --- AGGREGATE DATA BASED ON AGGREGATION LEVEL ---
    # Create a copy with Order Date as index
    df_sorted = df.sort_values("Order Date").copy()
    df_sorted.set_index("Order Date", inplace=True)

    agg_dict = {
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum",
        "Discount": "mean"
    }

    if agg_level == "Daily":
        df_agg = df_sorted.resample("D").agg(agg_dict).reset_index()
        rolling_window = 30
    elif agg_level == "Weekly":
        df_agg = df_sorted.resample("W").agg(agg_dict).reset_index()
        rolling_window = 4
    else:  # Monthly
        df_agg = df_sorted.resample("M").agg(agg_dict).reset_index()
        rolling_window = 3

    # Compute additional metrics
    df_agg["Margin Rate"] = df_agg["Profit"] / df_agg["Sales"].replace(0, 1)
    df_agg["Avg Discount"] = df_agg["Discount"]
    df_agg[f"Rolling_Avg_{selected_kpi}"] = df_agg[selected_kpi].rolling(window=rolling_window).mean()

    # --- 1. Time Series Area Chart with Rolling Average ---
    fig_area = px.area(
        df_agg,
        x="Order Date",
        y=selected_kpi,
        title=f"{selected_kpi} Over Time ({agg_level} Aggregation)",
        template=plotly_template,
        color_discrete_sequence=["#1E90FF"]
    )
    fig_area.add_scatter(
        x=df_agg["Order Date"],
        y=df_agg[f"Rolling_Avg_{selected_kpi}"],
        mode="lines+markers",
        name=f"{rolling_window} Period Rolling Avg",
        line=dict(color="orange")
    )
    fig_area.update_layout(
        hovermode="x unified",
        legend=dict(yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    # --- 2. Top 10 Products Chart ---
    product_grouped = df.groupby("Product Name", as_index=False).agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum",
        "Discount": "mean"
    })
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    product_grouped["Avg Discount"] = product_grouped["Discount"]
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)

    fig_bar = px.bar(
        top_10,
        x=selected_kpi,
        y="Product Name",
        orientation="h",
        title=f"Top 10 Products by {selected_kpi}",
        template=plotly_template,
        color=selected_kpi,
        color_continuous_scale="Blues"
    )
    fig_bar.update_layout(yaxis={"categoryorder": "total ascending"})

    # --- 3. Sales by Region as a Donut Chart ---
    region_grouped = df.groupby("Region", as_index=False)["Sales"].sum()
    fig_region = px.pie(
        region_grouped,
        names="Region",
        values="Sales",
        title="Sales by Region",
        template=plotly_template,
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    fig_region.update_traces(textposition='inside', textinfo='percent+label')

    # --- LAYOUT THE CHARTS ---
    col_left, col_right = st.columns(2)
    with col_left:
        st.plotly_chart(fig_area, use_container_width=True)
    with col_right:
        st.plotly_chart(fig_bar, use_container_width=True)

    col_region, col_download = st.columns(2)
    with col_region:
        st.plotly_chart(fig_region, use_container_width=True)
    with col_download:
        st.markdown("### Download Data")
        csv_data = df.reset_index(drop=True).to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download CSV",
            data=csv_data,
            file_name="filtered_superstore_data.csv",
            mime="text/csv"
        )
        st.dataframe(df.reset_index(drop=True), height=300, use_container_width=True)

    # ----- ADDITIONAL INSIGHTS -----
    st.markdown("## Additional Insights")
    # Sales & Profit Over Time: Dual-line chart using the aggregated data (using monthly for smoother trends)
    df_monthly = df_sorted.resample("M").agg(agg_dict).reset_index()
    df_monthly["Margin Rate"] = df_monthly["Profit"] / df_monthly["Sales"].replace(0, 1)
    fig_sales_profit = px.line(
        df_monthly,
        x="Order Date",
        y=["Sales", "Profit"],
        title="Sales and Profit Over Time (Monthly)",
        template=plotly_template
    )
    fig_sales_profit.update_layout(hovermode="x unified")

    # Profit by Category
    category_profit = df.groupby("Category", as_index=False)["Profit"].sum().sort_values(by="Profit", ascending=False)
    fig_profit_category = px.bar(
        category_profit,
        x="Category",
        y="Profit",
        title="Profit by Category",
        template=plotly_template,
        color="Profit",
        color_continuous_scale="Blues"
    )

    # Discount vs Profit Margin
    df["Profit Margin"] = df.apply(lambda row: row["Profit"] / row["Sales"] if row["Sales"] != 0 else 0, axis=1)
    fig_discount_profit = px.scatter(
        df,
        x="Discount",
        y="Profit Margin",
        color="Category",
        title="Discount vs Profit Margin",
        template=plotly_template,
        hover_data=["Sales", "Profit"]
    )

    # Sales Distribution by Sub-Category: Treemap
    subcat_sales = df.groupby("Sub-Category", as_index=False)["Sales"].sum()
    fig_treemap = px.treemap(
        subcat_sales,
        path=['Sub-Category'],
        values="Sales",
        title="Sales Distribution by Sub-Category",
        template=plotly_template,
        color="Sales",
        color_continuous_scale="Blues"
    )

    col_insight1, col_insight2 = st.columns(2)
    with col_insight1:
        st.plotly_chart(fig_sales_profit, use_container_width=True)
    with col_insight2:
        st.plotly_chart(fig_profit_category, use_container_width=True)

    col_insight3, col_insight4 = st.columns(2)
    with col_insight3:
        st.plotly_chart(fig_discount_profit, use_container_width=True)
    with col_insight4:
        st.plotly_chart(fig_treemap, use_container_width=True)
