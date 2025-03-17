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

# ---- Compute a Previous Period for Comparison ----
# Calculate the number of days in the current selection
selected_days = (pd.to_datetime(to_date) - pd.to_datetime(from_date)).days + 1
# Define the previous period as the same length of time immediately before the current period
prev_from_date = pd.to_datetime(from_date) - timedelta(days=selected_days)
prev_to_date = pd.to_datetime(from_date) - timedelta(days=1)

df_previous = df_filtered[
    (df_filtered["Order Date"] >= prev_from_date) &
    (df_filtered["Order Date"] <= prev_to_date)
]

# ---- Page Title ----
st.title("SuperStore KPI Dashboard")

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

# Calculate deltas (absolute differences)
sales_delta = current_sales - previous_sales
quantity_delta = current_quantity - previous_quantity
profit_delta = current_profit - previous_profit
margin_delta = (current_margin - previous_margin) * 100  # difference in margin points

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

# ---- KPI Display (Rectangles) with Deltas ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

with kpi_col1:
    # Sales
    if df_previous.empty:
        st.metric("Sales", f"${current_sales:,.2f}")
    else:
        st.metric("Sales", f"${current_sales:,.2f}", f"${sales_delta:,.2f}")

with kpi_col2:
    # Quantity Sold
    if df_previous.empty:
        st.metric("Quantity Sold", f"{current_quantity:,.0f}")
    else:
        st.metric("Quantity Sold", f"{current_quantity:,.0f}", f"{quantity_delta:,.0f}")

with kpi_col3:
    # Profit
    if df_previous.empty:
        st.metric("Profit", f"${current_profit:,.2f}")
    else:
        st.metric("Profit", f"${current_profit:,.2f}", f"${profit_delta:,.2f}")

with kpi_col4:
    # Margin Rate
    if df_previous.empty:
        st.metric("Margin Rate", f"{(current_margin * 100):,.2f}%")
    else:
        st.metric("Margin Rate", f"{(current_margin * 100):,.2f}%", f"{margin_delta:,.2f}%")

# ---- Sales Over Time (Line Chart) ----
st.subheader("Sales Over Time")
if df_current.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    daily_grouped = df_current.groupby("Order Date").agg({"Sales": "sum"}).reset_index()
    fig_line = px.line(
        daily_grouped,
        x="Order Date",
        y="Sales",
        title="Sales Over Time",
        labels={"Order Date": "Date", "Sales": "Sales"},
        template="plotly_white",
    )
    fig_line.update_layout(height=400)
    st.plotly_chart(fig_line, use_container_width=True)

# ---- Category & Sub-Category Distribution (Sunburst Chart) ----
st.subheader("Category & Sub-Category Sales Distribution")
if df_current.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    fig_sunburst = px.sunburst(
        df_current,
        path=["Category", "Sub-Category"],
        values="Sales",
        title="Category & Sub-Category Breakdown",
        template="plotly_white",
        color="Category",
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    fig_sunburst.update_layout(margin=dict(t=50, l=25, r=25, b=25))
    st.plotly_chart(fig_sunburst, use_container_width=True)

# ---- Top 10 Products by Selected KPI ----
st.subheader("Top 10 Products by Selected KPI")
if df_current.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
    selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)

    product_grouped = df_current.groupby("Product Name").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()

    # Calculate margin rate for each product
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)

    # Bar Chart
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
    fig_bar.update_layout(
        height=400,
        yaxis={"categoryorder": "total ascending"}
    )
    st.plotly_chart(fig_bar, use_container_width=True)
