"""
Sales Analytics Dashboard Runner
- Run with 'python chart.py' or 'python app.py': Launches the full-fledged FastAPI backend + modern interactive dashboard!
- Run with 'streamlit run chart.py': Launches the Streamlit dashboard interface!
"""
import sys

def is_running_under_streamlit() -> bool:
    """Check if executed via streamlit run."""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False

if __name__ == "__main__" and not is_running_under_streamlit():
    print("=" * 70)
    print("[SalesPulse Enterprise Dashboard - Full-Stack Backend + Frontend]")
    print("[*] Starting FastAPI & Uvicorn Server...")
    print("=" * 70)
    from app import start_server
    start_server()
    sys.exit(0)

import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# =============================================================================
# 1. PAGE CONFIG & STYLE
# =============================================================================
st.set_page_config(
    page_title="Sales Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        background-color: #1e293b;
        border: 1px solid #334155;
        padding: 14px 16px;
        border-radius: 12px;
    }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# 2. HELPERS
# =============================================================================
EXCEL_NAME = "sales.xlsx"
SHEET_NAME = "Sales_Data"
REQUIRED_COLUMNS = [
    "Order_Date", "Order_ID", "Region", "Category", "Product",
    "Customer_Name", "Quantity", "Discount_Percent", "Total_Amount_INR",
]
PLOT_COLORS = px.colors.qualitative.Bold


def fmt_inr(x: float) -> str:
    """Format a rupee amount as Cr / L / plain so it always fits in a card."""
    x = float(x)
    if abs(x) >= 1e7:
        return f"₹{x / 1e7:,.2f} Cr"
    if abs(x) >= 1e5:
        return f"₹{x / 1e5:,.2f} L"
    return f"₹{x:,.0f}"


def show_chart(fig):
    """Works on both old and new Streamlit versions."""
    try:
        st.plotly_chart(fig, width="stretch")
    except TypeError:
        st.plotly_chart(fig, use_container_width=True)


def show_table(data):
    try:
        st.dataframe(data, width="stretch", height=360)
    except TypeError:
        st.dataframe(data, use_container_width=True, height=360)


def style(fig, height=380):
    fig.update_layout(
        template="plotly_dark",
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        legend_title_text="",
    )
    return fig


@st.cache_data(show_spinner="Loading data...")
def load_data(source) -> pd.DataFrame:
    data = pd.read_excel(source, sheet_name=SHEET_NAME)
    data["Order_Date"] = pd.to_datetime(data["Order_Date"], errors="coerce")
    data = data.dropna(subset=["Order_Date"])
    data["Year_Month"] = data["Order_Date"].dt.to_period("M").astype(str)
    data["Year"] = data["Order_Date"].dt.year.astype(str)
    data["Weekday"] = data["Order_Date"].dt.day_name()
    return data


# =============================================================================
# 3. DATA LOADING (sales.xlsx next to this file, upload only as a backup)
# =============================================================================
FILE_PATH = Path(__file__).resolve().parent / EXCEL_NAME

try:
    if FILE_PATH.exists():
        df = load_data(FILE_PATH)
    else:
        st.warning(f"`{EXCEL_NAME}` was not found next to the app. Upload it below.")
        uploaded = st.file_uploader("Upload your Sales Excel file", type=["xlsx"])
        if uploaded is None:
            st.stop()
        df = load_data(uploaded)
except ValueError as err:
    st.error(f"Could not read the sheet `{SHEET_NAME}`: {err}")
    st.stop()

missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
if missing:
    st.error(f"These columns are missing from the Excel file: {', '.join(missing)}")
    st.stop()

# =============================================================================
# 4. SIDEBAR FILTERS
# =============================================================================
st.sidebar.header("🔍 Filter Dashboard")

min_date = df["Order_Date"].min().date()
max_date = df["Order_Date"].max().date()

date_range = st.sidebar.date_input(
    "Order date range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)

all_regions = sorted(df["Region"].dropna().unique())
selected_regions = st.sidebar.multiselect("Region", all_regions, default=all_regions)

all_categories = sorted(df["Category"].dropna().unique())
selected_categories = st.sidebar.multiselect("Category", all_categories, default=all_categories)

product_options = sorted(
    df[df["Category"].isin(selected_categories)]["Product"].dropna().unique()
)
selected_products = st.sidebar.multiselect("Product", product_options, default=product_options)

top_n = st.sidebar.slider("Top N products / customers", 3, 15, 10)

# The date picker returns 1 date while the user is still choosing the 2nd one
if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
else:
    start_date, end_date = pd.to_datetime(min_date), pd.to_datetime(max_date)

mask = (
    (df["Order_Date"] >= start_date)
    & (df["Order_Date"] < end_date + pd.Timedelta(days=1))
    & (df["Region"].isin(selected_regions))
    & (df["Category"].isin(selected_categories))
    & (df["Product"].isin(selected_products))
)
fdf = df[mask]

# =============================================================================
# 5. HEADER & KPIs
# =============================================================================
st.title("📊 Enterprise Sales Analytics Dashboard")

if fdf.empty:
    st.warning("No data matches the current filters. Widen the filters in the sidebar.")
    st.stop()

st.caption(
    f"Showing **{len(fdf):,}** orders from **{fdf['Order_Date'].min().date()}** "
    f"to **{fdf['Order_Date'].max().date()}**"
)

total_revenue = fdf["Total_Amount_INR"].sum()
total_orders = fdf["Order_ID"].nunique()
total_units = int(fdf["Quantity"].sum())
avg_order_value = total_revenue / total_orders if total_orders else 0
avg_discount = fdf["Discount_Percent"].mean()

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Revenue", fmt_inr(total_revenue))
k2.metric("Total Orders", f"{total_orders:,}")
k3.metric("Units Sold", f"{total_units:,}")
k4.metric("Avg Order Value", fmt_inr(avg_order_value))
k5.metric("Avg Discount", f"{avg_discount:.1f}%")

st.markdown("---")

# =============================================================================
# 6. TABS
# =============================================================================
tab_overview, tab_products, tab_insights, tab_data = st.tabs(
    ["📈 Overview", "🏆 Products & Regions", "💡 Insights", "📋 Data"]
)

# ---------------------------------- Overview ---------------------------------
with tab_overview:
    c1, c2 = st.columns([2, 1])

    with c1:
        st.subheader("Monthly Revenue Trend")
        monthly = (
            fdf.groupby("Year_Month", as_index=False)["Total_Amount_INR"]
            .sum()
            .sort_values("Year_Month")
        )
        fig = px.area(
            monthly, x="Year_Month", y="Total_Amount_INR",
            labels={"Year_Month": "Month", "Total_Amount_INR": "Revenue (₹)"},
            color_discrete_sequence=["#38bdf8"],
        )
        show_chart(style(fig))

    with c2:
        st.subheader("Revenue by Category")
        cat = fdf.groupby("Category", as_index=False)["Total_Amount_INR"].sum()
        fig = px.pie(
            cat, names="Category", values="Total_Amount_INR", hole=0.5,
            color_discrete_sequence=PLOT_COLORS,
        )
        fig.update_traces(textposition="inside", textinfo="percent")
        show_chart(style(fig))

    c3, c4 = st.columns(2)

    with c3:
        st.subheader("Revenue by Year")
        yearly = fdf.groupby("Year", as_index=False)["Total_Amount_INR"].sum()
        fig = px.bar(
            yearly, x="Year", y="Total_Amount_INR", text_auto=".3s",
            labels={"Total_Amount_INR": "Revenue (₹)"},
            color_discrete_sequence=["#a78bfa"],
        )
        show_chart(style(fig, 340))

    with c4:
        st.subheader("Orders by Weekday")
        order_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        wk = fdf.groupby("Weekday", as_index=False)["Order_ID"].nunique()
        wk["Weekday"] = pd.Categorical(wk["Weekday"], categories=order_days, ordered=True)
        wk = wk.sort_values("Weekday")
        fig = px.bar(
            wk, x="Weekday", y="Order_ID", text_auto=True,
            labels={"Order_ID": "Orders"},
            color_discrete_sequence=["#34d399"],
        )
        show_chart(style(fig, 340))

# ----------------------------- Products & Regions ----------------------------
with tab_products:
    p1, p2 = st.columns(2)

    with p1:
        st.subheader(f"Top {top_n} Products by Revenue")
        prod = (
            fdf.groupby(["Product", "Category"], as_index=False)["Total_Amount_INR"]
            .sum()
            .sort_values("Total_Amount_INR", ascending=False)
            .head(top_n)
            .sort_values("Total_Amount_INR", ascending=True)
        )
        fig = px.bar(
            prod, x="Total_Amount_INR", y="Product", color="Category", orientation="h",
            labels={"Total_Amount_INR": "Revenue (₹)"},
            color_discrete_sequence=PLOT_COLORS,
        )
        show_chart(style(fig, 440))

    with p2:
        st.subheader("Regional Sales Performance")
        reg = fdf.groupby(["Region", "Category"], as_index=False)["Total_Amount_INR"].sum()
        fig = px.bar(
            reg, x="Region", y="Total_Amount_INR", color="Category", barmode="group",
            labels={"Total_Amount_INR": "Revenue (₹)"},
            color_discrete_sequence=PLOT_COLORS,
        )
        show_chart(style(fig, 440))

    st.subheader("Region × Category Heatmap (Revenue)")
    heat = fdf.pivot_table(
        index="Region", columns="Category", values="Total_Amount_INR",
        aggfunc="sum", fill_value=0,
    )
    fig = px.imshow(
        heat, text_auto=".2s", aspect="auto", color_continuous_scale="Blues",
        labels=dict(color="Revenue (₹)"),
    )
    show_chart(style(fig, 360))

# --------------------------------- Insights ----------------------------------
with tab_insights:
    best_region = fdf.groupby("Region")["Total_Amount_INR"].sum().idxmax()
    best_category = fdf.groupby("Category")["Total_Amount_INR"].sum().idxmax()
    best_product = fdf.groupby("Product")["Total_Amount_INR"].sum().idxmax()
    best_month = fdf.groupby("Year_Month")["Total_Amount_INR"].sum().idxmax()

    i1, i2, i3, i4 = st.columns(4)
    i1.metric("Top Region", best_region)
    i2.metric("Top Category", best_category)
    i3.metric("Top Product", best_product)
    i4.metric("Best Month", best_month)

    s1, s2 = st.columns(2)

    with s1:
        st.subheader(f"Top {top_n} Customers by Revenue")
        cust = (
            fdf.groupby("Customer_Name", as_index=False)["Total_Amount_INR"]
            .sum()
            .sort_values("Total_Amount_INR", ascending=False)
            .head(top_n)
            .sort_values("Total_Amount_INR", ascending=True)
        )
        fig = px.bar(
            cust, x="Total_Amount_INR", y="Customer_Name", orientation="h",
            labels={"Total_Amount_INR": "Revenue (₹)", "Customer_Name": "Customer"},
            color_discrete_sequence=["#f472b6"],
        )
        show_chart(style(fig, 440))

    with s2:
        st.subheader("Average Discount by Category")
        disc = fdf.groupby("Category", as_index=False)["Discount_Percent"].mean()
        fig = px.bar(
            disc, x="Category", y="Discount_Percent", text_auto=".1f",
            labels={"Discount_Percent": "Avg Discount (%)"},
            color_discrete_sequence=["#fbbf24"],
        )
        show_chart(style(fig, 440))

# ----------------------------------- Data ------------------------------------
with tab_data:
    st.subheader("Filtered Order Records & Export")
    query = st.text_input("🔎 Search by Customer Name or Order ID", "")

    view = fdf
    if query:
        view = view[
            view["Customer_Name"].astype(str).str.contains(query, case=False, na=False, regex=False)
            | view["Order_ID"].astype(str).str.contains(query, case=False, na=False, regex=False)
        ]

    export_df = view.drop(columns=["Year_Month", "Year", "Weekday"])
    st.caption(f"{len(export_df):,} matching rows (table shows the first 500)")
    show_table(export_df.head(500))

    st.download_button(
        label="⬇️ Download Filtered Data as CSV",
        data=export_df.to_csv(index=False).encode("utf-8"),
        file_name="filtered_sales_data.csv",
        mime="text/csv",
    )