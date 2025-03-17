import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# Set page config for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ---- Sidebar Filters ----
st.sidebar.title("Filters")

all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)
if selected_region != "All":
    df_filtered_region = df_original[df_original["Region"] == selected_region]
else:
    df_filtered_region = df_original

all_states = sorted(df_filtered_region["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)
if selected_state != "All":
    df_filtered_state = df_filtered_region[df_filtered_region["State"] == selected_state]
else:
    df_filtered_state = df_filtered_region

all_categories = sorted(df_filtered_state["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)
if selected_category != "All":
    df_filtered_category = df_filtered_state[df_filtered_state["Category"] == selected_category]
else:
    df_filtered_category = df_filtered_state

all_subcats = sorted(df_filtered_category["Sub-Category"].dropna().unique())
selected_subcat = st.sidebar.selectbox("Select Sub-Category", options=["All"] + all_subcats)
df_filtered = df_filtered_category.copy()
if selected_subcat != "All":
    df_filtered = df_filtered[df_filtered["Sub-Category"] == selected_subcat]

# ---- Sidebar Date Range ----
if df_filtered.empty:
    min_date = df_original["Order Date"].min()
    max_date = df_original["Order Date"].max()
else:
    min_date = df_filtered["Order Date"].min()
    max_date = df_filtered["Order Date"].max()

from_date = st.sidebar.date_input("From Date", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("To Date", value=max_date, min_value=min_date, max_value=max_date)

if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

df_current = df_filtered[
    (df_filtered["Order Date"] >= pd.to_datetime(from_date)) &
    (df_filtered["Order Date"] <= pd.to_datetime(to_date))
]

# ---- Compute Previous Period for Comparison ----
selected_days = (pd.to_datetime(to_date) - pd.to_datetime(from_date)).days + 1
prev_from_date = pd.to_datetime(from_date) - timedelta(days=selected_days)
prev_to_date = pd.to_datetime(from_date) - timedelta(days=1)
df_previous = df_filtered[
    (df_filtered["Order Date"] >= prev_from_date) &
    (df_filtered["Order Date"] <= prev_to_date)
]

# ---- Calculate Current & Previous Metrics ----
def safe_sum(df, col):
    return df[col].sum() if not df.empty else 0

current_sales = safe_sum(df_current, "Sales")
current_quantity = safe_sum(df_current, "Quantity")
current_profit = safe_sum(df_current, "Profit")
current_margin = (current_profit / current_sales) if current_sales != 0 else 0

previous_sales = safe_sum(df_previous, "Sales")
previous_quantity = safe_sum(df_previous, "Quantity")
previous_profit = safe_sum(df_previous, "Profit")
previous_margin = (previous_profit / previous_sales) if previous_sales != 0 else 0

sales_delta = current_sales - previous_sales
quantity_delta = current_quantity - previous_quantity
profit_delta = current_profit - previous_profit
margin_delta = (current_margin - previous_margin) * 100  # in percentage points

# ---- Page Title ----
st.title("SuperStore KPI Dashboard")

# ---- Custom CSS for KPI Tiles ----
st.markdown(
    """
    <style>
    .kpi-box {
        background-color: #FFFFFF;
        border: 2px solid #EAEAEA;
        border-radius: 8px;
        padding: 16px;
        margin: 8px;
        text-align: center;
    }
    .kpi-title {
        font-weight: 600;
        color: #333333;
        font-size: 16px;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-weight: 700;
        font-size: 24px;
        color: #1E90FF;
    }
    .kpi-delta {
        font-size: 14px;
        color: #555555;
        margin-top: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---- KPI Display (Boxes with Delta Indicators) ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
with kpi_col1:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Sales</div>
            <div class='kpi-value'>${current_sales:,.2f}</div>
            <div class='kpi-delta'>Change: ${sales_delta:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col2:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Quantity Sold</div>
            <div class='kpi-value'>{current_quantity:,.0f}</div>
            <div class='kpi-delta'>Change: {quantity_delta:,.0f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col3:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Profit</div>
            <div class='kpi-value'>${current_profit:,.2f}</div>
            <div class='kpi-delta'>Change: ${profit_delta:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col4:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Margin Rate</div>
            <div class='kpi-value'>{(current_margin * 100):,.2f}%</div>
            <div class='kpi-delta'>Change: {margin_delta:,.2f}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---- KPI Toggle for Charts ----
st.subheader("Visualize KPI Across Time, Categories, & Top Products")
kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)

# ---- Chart Type Toggle for Time Series ----
chart_type = st.radio("Select Chart Type for Time Series", options=["Line Chart", "Area Chart"], horizontal=True)

# ---- Time Series Chart for Selected KPI ----
if df_current.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    daily_grouped = df_current.groupby("Order Date").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)
    
    if chart_type == "Line Chart":
        fig_time = px.line(
            daily_grouped,
            x="Order Date",
            y=selected_kpi,
            title=f"{selected_kpi} Over Time",
            labels={"Order Date": "Date", selected_kpi: selected_kpi},
            template="plotly_white",
        )
    else:
        fig_time = px.area(
            daily_grouped,
            x="Order Date",
            y=selected_kpi,
            title=f"{selected_kpi} Over Time",
            labels={"Order Date": "Date", selected_kpi: selected_kpi},
            template="plotly_white",
        )
    fig_time.update_layout(height=400)
    st.plotly_chart(fig_time, use_container_width=True)

# ---- Category Breakdown Chart for Selected KPI (Histogram) ----
if df_current.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    category_grouped = df_current.groupby("Category").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    category_grouped["Margin Rate"] = category_grouped["Profit"] / category_grouped["Sales"].replace(0, 1)
    
    fig_category = px.bar(
        category_grouped,
        x="Category",
        y=selected_kpi,
        title=f"{selected_kpi} by Category",
        labels={"Category": "Category", selected_kpi: selected_kpi},
        template="plotly_white",
        color="Category"
    )
    fig_category.update_layout(height=400)
    st.plotly_chart(fig_category, use_container_width=True)

# ---- Top 10 Products by Selected KPI ----
st.subheader(f"Top 10 Products by {selected_kpi}")
if df_current.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    product_grouped = df_current.groupby("Product Name").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)
    
    fig_top10 = px.bar(
        top_10,
        x=selected_kpi,
        y="Product Name",
        orientation="h",
        title=f"Top 10 Products by {selected_kpi}",
        labels={selected_kpi: selected_kpi, "Product Name": "Product"},
        color=selected_kpi,
        color_continuous_scale="Blues",
        template="plotly_white",
    )
    fig_top10.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_top10, use_container_width=True)
