import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# Set page config for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    # Adjust the path if needed, e.g. "data/Sample - Superstore.xlsx"
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    # Convert Order Date to datetime if not already
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ---- Sidebar Filters ----
st.sidebar.title("Filters")

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)
if selected_region != "All":
    df_filtered_region = df_original[df_original["Region"] == selected_region]
else:
    df_filtered_region = df_original

# State Filter
all_states = sorted(df_filtered_region["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)
if selected_state != "All":
    df_filtered_state = df_filtered_region[df_filtered_region["State"] == selected_state]
else:
    df_filtered_state = df_filtered_region

# Category Filter
all_categories = sorted(df_filtered_state["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)
if selected_category != "All":
    df_filtered_category = df_filtered_state[df_filtered_state["Category"] == selected_category]
else:
    df_filtered_category = df_filtered_state

# Sub-Category Filter
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

# Apply date range filter to current data
df_current = df_filtered[
    (df_filtered["Order Date"] >= pd.to_datetime(from_date)) &
    (df_filtered["Order Date"] <= pd.to_datetime(to_date))
]

# ---- Compute Previous Period for Performance Trends ----
selected_days = (pd.to_datetime(to_date) - pd.to_datetime(from_date)).days + 1
prev_from_date = pd.to_datetime(from_date) - timedelta(days=selected_days)
prev_to_date = pd.to_datetime(from_date) - timedelta(days=1)
df_previous = df_filtered[
    (df_filtered["Order Date"] >= prev_from_date) &
    (df_filtered["Order Date"] <= prev_to_date)
]

# ---- Page Title ----
st.title("SuperStore KPI Dashboard")

# ---- Performance Trends Section (New Enhancement) ----
st.markdown("## Performance Trends")
if df_current.empty:
    st.write("No current period data available for the selected filters and date range.")
else:
    # Current period metrics
    current_sales = df_current["Sales"].sum()
    current_profit = df_current["Profit"].sum()
    current_margin = (current_profit / current_sales) if current_sales != 0 else 0

    # Previous period metrics
    previous_sales = df_previous["Sales"].sum() if not df_previous.empty else 0
    previous_profit = df_previous["Profit"].sum() if not df_previous.empty else 0
    previous_margin = (previous_profit / previous_sales) if previous_sales != 0 else 0

    # Function to calculate percentage change
    def pct_change(current, previous):
        return ((current - previous) / previous * 100) if previous != 0 else float('inf')

    sales_change = pct_change(current_sales, previous_sales)
    profit_change = pct_change(current_profit, previous_profit)
    margin_change = pct_change(current_margin, previous_margin)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Sales", f"${current_sales:,.2f}", f"{sales_change:,.2f}%")
    with col2:
        st.metric("Total Profit", f"${current_profit:,.2f}", f"{profit_change:,.2f}%")
    with col3:
        st.metric("Margin Rate", f"{(current_margin * 100):,.2f}%", f"{margin_change:,.2f}%")

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
    </style>
    """,
    unsafe_allow_html=True
)

# ---- KPI Calculation for Main Display ----
if df_current.empty:
    total_sales = 0
    total_quantity = 0
    total_profit = 0
    margin_rate = 0
else:
    total_sales = df_current["Sales"].sum()
    total_quantity = df_current["Quantity"].sum()
    total_profit = df_current["Profit"].sum()
    margin_rate = (total_profit / total_sales) if total_sales != 0 else 0

# ---- KPI Display (Rectangles) ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
with kpi_col1:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Sales</div>
            <div class='kpi-value'>${total_sales:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col2:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Quantity Sold</div>
            <div class='kpi-value'>{total_quantity:,.0f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col3:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Profit</div>
            <div class='kpi-value'>${total_profit:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col4:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Margin Rate</div>
            <div class='kpi-value'>{(margin_rate * 100):,.2f}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---- KPI Selection (Affects Both Charts) ----
st.subheader("Visualize KPI Across Time & Top Products")

if df_current.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
    selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)
    
    # Chart Type Toggle for Time Series (Line or Area Chart)
    chart_type = st.radio("Select Chart Type for Time Series", options=["Line Chart", "Area Chart"], horizontal=True)
    
    # ---- Prepare Data for Charts ----
    daily_grouped = df_current.groupby("Order Date").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)
    
    product_grouped = df_current.groupby("Product Name").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)
    
    col_left, col_right = st.columns(2)
    
    with col_left:
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
    
    with col_right:
        fig_bar = px.bar(
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
        fig_bar.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_bar, use_container_width=True)
