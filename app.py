"""
Sales Analytics Dashboard - Enterprise Web Application
Full-Stack Architecture:
  - Backend: FastAPI, Uvicorn, Pandas, OpenpyXL (RESTful API & vectorized analytics)
  - Frontend: Responsive SPA with Tailwind CSS, Chart.js, Lucide Icons, Vanilla JS

Usage:
  python app.py
  (Opens http://127.0.0.1:8000 automatically in your browser)
"""

import io
import math
import os
import sys
import threading
import time
import webbrowser
from pathlib import Path
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
import uvicorn
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response, StreamingResponse

# =============================================================================
# 1. CONSTANTS & PATHS
# =============================================================================
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL_FILE = BASE_DIR / "sales.xlsx"
CACHE_FILE = BASE_DIR / ".sales_cache.pkl"
DEFAULT_SHEET_NAME = "Sales_Data"
TEMPLATES_DIR = BASE_DIR / "templates"
INDEX_HTML_PATH = TEMPLATES_DIR / "index.html"

REQUIRED_COLUMNS = [
    "Order_Date",
    "Order_ID",
    "Region",
    "Category",
    "Product",
    "Customer_Name",
    "Quantity",
    "Discount_Percent",
    "Total_Amount_INR",
]

# Order of days for weekday charts
WEEKDAY_ORDER = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]


# =============================================================================
# 2. DATA MANAGER (BACKEND DATA ACCESS LAYER)
# =============================================================================
class DataManager:
    """Manages the in-memory dataset, vectorized pandas caching, and analytical queries."""

    def __init__(self, file_path: Path = DEFAULT_EXCEL_FILE):
        self.file_path = file_path
        self.df: pd.DataFrame = pd.DataFrame()
        self.last_loaded: float = 0.0
        self.filename: str = file_path.name
        self._lock = threading.Lock()
        self.load_dataset()

    def load_dataset(self, source=None, filename: Optional[str] = None):
        """Load and preprocess sales data into memory."""
        with self._lock:
            target = source if source is not None else self.file_path
            name = filename or (self.file_path.name if isinstance(target, Path) else "uploaded_data")

            if isinstance(target, Path) and not target.exists():
                print(f"[Backend Warning] Data file '{target}' does not exist.")
                self.df = pd.DataFrame(columns=REQUIRED_COLUMNS)
                self.filename = name
                self.last_loaded = time.time()
                return

            print(f"[Backend] Loading dataset from '{name}'...")
            start_t = time.time()

            # Fast cache hit for default dataset
            if target == self.file_path and CACHE_FILE.exists():
                try:
                    if CACHE_FILE.stat().st_mtime >= self.file_path.stat().st_mtime:
                        df = pd.read_pickle(CACHE_FILE)
                        self.df = df
                        self.filename = name
                        self.last_loaded = time.time()
                        elapsed = round(time.time() - start_t, 3)
                        print(f"[Backend] Loaded {len(self.df):,} rows from fast cache in {elapsed}s.")
                        return
                except Exception as ce:
                    print(f"[Backend Cache Notice] Re-parsing Excel (cache invalid: {ce})")

            try:
                if str(name).lower().endswith(".csv"):
                    df = pd.read_csv(target)
                else:
                    # Excel file
                    try:
                        df = pd.read_excel(target, sheet_name=DEFAULT_SHEET_NAME)
                    except Exception:
                        df = pd.read_excel(target)
            except Exception as e:
                print(f"[Backend Error] Failed reading file: {e}")
                raise ValueError(f"Could not load data file: {e}")

            # Validate columns
            missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
            if missing:
                err_msg = f"Missing required columns in dataset: {', '.join(missing)}"
                print(f"[Backend Error] {err_msg}")
                raise ValueError(err_msg)

            # Preprocess dates
            df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
            df = df.dropna(subset=["Order_Date"])

            # Feature Engineering for fast aggregations
            df["Order_Date_Str"] = df["Order_Date"].dt.strftime("%Y-%m-%d")
            df["Year_Month"] = df["Order_Date"].dt.strftime("%Y-%m")
            df["Year"] = df["Order_Date"].dt.year.astype(str)
            df["Weekday"] = df["Order_Date"].dt.day_name()

            # Numeric cleaning
            df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(1).astype(int)
            df["Total_Amount_INR"] = pd.to_numeric(df["Total_Amount_INR"], errors="coerce").fillna(0.0)
            df["Discount_Percent"] = pd.to_numeric(df["Discount_Percent"], errors="coerce").fillna(0.0)
            if "Unit_Price_INR" in df.columns:
                df["Unit_Price_INR"] = pd.to_numeric(df["Unit_Price_INR"], errors="coerce").fillna(0.0)
            else:
                df["Unit_Price_INR"] = (df["Total_Amount_INR"] / df["Quantity"]).round(2)

            # Calculate total discount savings
            gross = df["Quantity"] * df["Unit_Price_INR"]
            df["Discount_Amount_INR"] = (gross - df["Total_Amount_INR"]).clip(lower=0.0)

            # Cache default dataset
            if target == self.file_path:
                try:
                    df.to_pickle(CACHE_FILE)
                except Exception:
                    pass

            self.df = df
            self.filename = name
            self.last_loaded = time.time()
            elapsed = round(time.time() - start_t, 2)
            print(f"[Backend] Successfully loaded {len(self.df):,} rows in {elapsed}s.")

    def filter_df(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        regions: Optional[List[str]] = None,
        categories: Optional[List[str]] = None,
        product: Optional[str] = None,
    ) -> pd.DataFrame:
        """Vectorized filter query based on parameters."""
        if self.df.empty:
            return self.df

        mask = pd.Series(True, index=self.df.index)

        if start_date:
            try:
                s_dt = pd.to_datetime(start_date)
                mask &= self.df["Order_Date"] >= s_dt
            except Exception:
                pass

        if end_date:
            try:
                e_dt = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(nanoseconds=1)
                mask &= self.df["Order_Date"] <= e_dt
            except Exception:
                pass

        if regions and len(regions) > 0:
            mask &= self.df["Region"].isin(regions)

        if categories and len(categories) > 0:
            mask &= self.df["Category"].isin(categories)

        if product and product.strip():
            mask &= self.df["Product"] == product.strip()

        return self.df[mask]

    def get_filter_options(self) -> Dict[str, Any]:
        """Return available unique options for dashboard filters."""
        if self.df.empty:
            return {
                "min_date": "",
                "max_date": "",
                "regions": [],
                "categories": [],
                "products_by_category": {},
                "all_products": [],
                "total_records": 0,
            }

        min_d = self.df["Order_Date"].min().strftime("%Y-%m-%d")
        max_d = self.df["Order_Date"].max().strftime("%Y-%m-%d")
        regions = sorted(self.df["Region"].dropna().unique().tolist())
        categories = sorted(self.df["Category"].dropna().unique().tolist())
        all_products = sorted(self.df["Product"].dropna().unique().tolist())

        # Products grouped by category
        prod_cat = {}
        for cat, group in self.df.groupby("Category"):
            prod_cat[str(cat)] = sorted(group["Product"].dropna().unique().tolist())

        return {
            "min_date": min_d,
            "max_date": max_d,
            "regions": regions,
            "categories": categories,
            "products_by_category": prod_cat,
            "all_products": all_products,
            "total_records": int(len(self.df)),
        }

    @staticmethod
    def format_inr(x: float) -> str:
        """Format INR currency into friendly Cr / L / K denominations."""
        x = float(x)
        if abs(x) >= 1e7:
            return f"₹{x / 1e7:,.2f} Cr"
        if abs(x) >= 1e5:
            return f"₹{x / 1e5:,.2f} L"
        if abs(x) >= 1e3:
            return f"₹{x / 1e3:,.1f} K"
        return f"₹{x:,.0f}"

    def get_kpis(self, fdf: pd.DataFrame) -> Dict[str, Any]:
        """Compute key business metrics on filtered DataFrame."""
        total_records = len(self.df)
        filtered_count = len(fdf)

        if fdf.empty:
            return {
                "total_revenue": 0.0,
                "formatted_revenue": "₹0",
                "total_orders": 0,
                "total_units": 0,
                "avg_order_value": 0.0,
                "formatted_aov": "₹0",
                "avg_discount": 0.0,
                "total_discount_amount": 0.0,
                "filtered_count": 0,
                "total_count": total_records,
            }

        total_rev = float(fdf["Total_Amount_INR"].sum())
        total_ord = int(fdf["Order_ID"].nunique())
        total_units = int(fdf["Quantity"].sum())
        aov = (total_rev / total_ord) if total_ord > 0 else 0.0
        avg_disc = float(fdf["Discount_Percent"].mean())
        total_disc_amt = float(fdf["Discount_Amount_INR"].sum())

        return {
            "total_revenue": round(total_rev, 2),
            "formatted_revenue": self.format_inr(total_rev),
            "total_orders": total_ord,
            "total_units": total_units,
            "avg_order_value": round(aov, 2),
            "formatted_aov": self.format_inr(aov),
            "avg_discount": round(avg_disc, 2),
            "total_discount_amount": round(total_disc_amt, 2),
            "filtered_count": filtered_count,
            "total_count": total_records,
        }

    def get_trend(self, fdf: pd.DataFrame) -> List[Dict[str, Any]]:
        """Time-series monthly trend of revenue, order volume, and units."""
        if fdf.empty:
            return []

        grouped = (
            fdf.groupby("Year_Month")
            .agg(
                revenue=("Total_Amount_INR", "sum"),
                orders=("Order_ID", "nunique"),
                units=("Quantity", "sum"),
            )
            .reset_index()
            .sort_values("Year_Month")
        )

        return [
            {
                "period": str(row["Year_Month"]),
                "revenue": round(float(row["revenue"]), 2),
                "orders": int(row["orders"]),
                "units": int(row["units"]),
            }
            for _, row in grouped.iterrows()
        ]

    def get_category_distribution(self, fdf: pd.DataFrame) -> List[Dict[str, Any]]:
        """Category-wise revenue distribution with percentages."""
        if fdf.empty:
            return []

        cat_grouped = (
            fdf.groupby("Category")
            .agg(
                revenue=("Total_Amount_INR", "sum"),
                orders=("Order_ID", "nunique"),
                units=("Quantity", "sum"),
            )
            .reset_index()
            .sort_values("revenue", ascending=False)
        )

        total_revenue = cat_grouped["revenue"].sum()
        result = []
        for _, row in cat_grouped.iterrows():
            rev = float(row["revenue"])
            share = (rev / total_revenue * 100) if total_revenue > 0 else 0.0
            result.append(
                {
                    "category": str(row["Category"]),
                    "revenue": round(rev, 2),
                    "orders": int(row["orders"]),
                    "units": int(row["units"]),
                    "percentage": round(share, 1),
                }
            )
        return result

    def get_regional_breakdown(self, fdf: pd.DataFrame) -> Dict[str, Any]:
        """Cross-tabulation of Revenue by Region and Category."""
        if fdf.empty:
            return {"regions": [], "categories": [], "matrix": {}}

        pivot = fdf.pivot_table(
            index="Region",
            columns="Category",
            values="Total_Amount_INR",
            aggfunc="sum",
            fill_value=0.0,
        )

        regions = pivot.index.astype(str).tolist()
        categories = pivot.columns.astype(str).tolist()
        matrix = {}
        for reg in regions:
            matrix[reg] = {}
            for cat in categories:
                matrix[reg][cat] = round(float(pivot.loc[reg, cat]), 2)

        return {"regions": regions, "categories": categories, "matrix": matrix}

    def get_weekday_dynamics(self, fdf: pd.DataFrame) -> List[Dict[str, Any]]:
        """Sales volume and order distribution by day of the week."""
        if fdf.empty:
            return []

        grouped = (
            fdf.groupby("Weekday")
            .agg(orders=("Order_ID", "nunique"), revenue=("Total_Amount_INR", "sum"))
            .reindex(WEEKDAY_ORDER)
            .fillna(0)
            .reset_index()
        )

        return [
            {
                "day": str(row["Weekday"]),
                "orders": int(row["orders"]),
                "revenue": round(float(row["revenue"]), 2),
            }
            for _, row in grouped.iterrows()
        ]

    def get_top_products(self, fdf: pd.DataFrame, limit: int = 10, metric: str = "revenue") -> List[Dict[str, Any]]:
        """Top products sorted by revenue or quantity."""
        if fdf.empty:
            return []

        grouped = (
            fdf.groupby(["Product", "Category"])
            .agg(
                revenue=("Total_Amount_INR", "sum"),
                units=("Quantity", "sum"),
                orders=("Order_ID", "nunique"),
            )
            .reset_index()
        )

        sort_col = "revenue" if metric == "revenue" else "units"
        top = grouped.sort_values(sort_col, ascending=False).head(limit)

        return [
            {
                "product": str(row["Product"]),
                "category": str(row["Category"]),
                "revenue": round(float(row["revenue"]), 2),
                "units": int(row["units"]),
                "orders": int(row["orders"]),
            }
            for _, row in top.iterrows()
        ]

    def get_top_customers(self, fdf: pd.DataFrame, limit: int = 10) -> List[Dict[str, Any]]:
        """Top VIP customers ranked by total spend."""
        if fdf.empty:
            return []

        top = (
            fdf.groupby("Customer_Name")
            .agg(
                revenue=("Total_Amount_INR", "sum"),
                orders=("Order_ID", "nunique"),
                units=("Quantity", "sum"),
            )
            .reset_index()
            .sort_values("revenue", ascending=False)
            .head(limit)
        )

        return [
            {
                "customer_name": str(row["Customer_Name"]),
                "revenue": round(float(row["revenue"]), 2),
                "orders": int(row["orders"]),
                "units": int(row["units"]),
            }
            for _, row in top.iterrows()
        ]

    def get_executive_insights(self, fdf: pd.DataFrame) -> Dict[str, Any]:
        """Automated business intelligence superlatives."""
        if fdf.empty:
            return {
                "best_region": "None",
                "best_category": "None",
                "best_product": "None",
                "best_month": "None",
                "highest_order": None,
            }

        best_reg = fdf.groupby("Region")["Total_Amount_INR"].sum().idxmax()
        best_cat = fdf.groupby("Category")["Total_Amount_INR"].sum().idxmax()
        best_prod = fdf.groupby("Product")["Total_Amount_INR"].sum().idxmax()
        best_mo = fdf.groupby("Year_Month")["Total_Amount_INR"].sum().idxmax()

        # Highest single transaction
        max_idx = fdf["Total_Amount_INR"].idxmax()
        max_row = fdf.loc[max_idx]

        return {
            "best_region": str(best_reg),
            "best_category": str(best_cat),
            "best_product": str(best_prod),
            "best_month": str(best_mo),
            "highest_order": {
                "order_id": str(max_row["Order_ID"]),
                "customer": str(max_row["Customer_Name"]),
                "product": str(max_row["Product"]),
                "amount": round(float(max_row["Total_Amount_INR"]), 2),
            },
        }

    def get_paginated_orders(
        self,
        fdf: pd.DataFrame,
        search: Optional[str] = None,
        sort_by: str = "Order_Date",
        sort_dir: str = "desc",
        page: int = 1,
        page_size: int = 20,
    ) -> Dict[str, Any]:
        """Searchable, sortable, paginated transaction records."""
        if fdf.empty:
            return {"total": 0, "page": page, "page_size": page_size, "total_pages": 0, "orders": []}

        view = fdf

        # Live search across Customer_Name, Order_ID, Product
        if search and search.strip():
            s = search.strip().lower()
            mask = (
                view["Customer_Name"].astype(str).str.lower().str.contains(s, regex=False)
                | view["Order_ID"].astype(str).str.lower().str.contains(s, regex=False)
                | view["Product"].astype(str).str.lower().str.contains(s, regex=False)
            )
            view = view[mask]

        total_records = len(view)
        total_pages = math.ceil(total_records / page_size) if total_records > 0 else 1
        page = max(1, min(page, total_pages))

        # Sort
        valid_cols = [
            "Order_ID",
            "Order_Date",
            "Customer_Name",
            "Product",
            "Category",
            "Region",
            "Quantity",
            "Unit_Price_INR",
            "Discount_Percent",
            "Total_Amount_INR",
        ]
        target_sort = sort_by if sort_by in valid_cols else "Order_Date"
        ascending = sort_dir.lower() == "asc"

        view_sorted = view.sort_values(by=target_sort, ascending=ascending)

        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        slice_df = view_sorted.iloc[start_idx:end_idx]

        orders = []
        for _, r in slice_df.iterrows():
            orders.append(
                {
                    "Order_ID": str(r["Order_ID"]),
                    "Order_Date": str(r["Order_Date_Str"]),
                    "Customer_Name": str(r["Customer_Name"]),
                    "Product": str(r["Product"]),
                    "Category": str(r["Category"]),
                    "Region": str(r["Region"]),
                    "Quantity": int(r["Quantity"]),
                    "Unit_Price_INR": round(float(r["Unit_Price_INR"]), 2),
                    "Discount_Percent": round(float(r["Discount_Percent"]), 1),
                    "Total_Amount_INR": round(float(r["Total_Amount_INR"]), 2),
                }
            )

        return {
            "total": total_records,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages,
            "orders": orders,
        }

    def export_csv(self, fdf: pd.DataFrame) -> io.StringIO:
        """Export current filtered view to CSV format."""
        export_cols = [
            "Order_ID",
            "Order_Date_Str",
            "Customer_Name",
            "Product",
            "Category",
            "Region",
            "Quantity",
            "Unit_Price_INR",
            "Discount_Percent",
            "Total_Amount_INR",
        ]
        export_df = fdf[[c for c in export_cols if c in fdf.columns]].rename(
            columns={"Order_Date_Str": "Order_Date"}
        )
        buffer = io.StringIO()
        export_df.to_csv(buffer, index=False)
        buffer.seek(0)
        return buffer


# Initialize Singleton Data Manager
db = DataManager()


# =============================================================================
# 3. FASTAPI BACKEND APPLICATION
# =============================================================================
app = FastAPI(
    title="SalesPulse Enterprise Analytics API",
    description="Production-grade Full-Stack Sales Analytics Dashboard REST API",
    version="2.0.0",
)

# Enable CORS for universal frontend compatibility
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def clean_str(val: Any, default: Optional[str] = None) -> Optional[str]:
    """Helper to ensure clean string when called directly or via Query."""
    if val is None or not isinstance(val, str):
        return default
    return val.strip() if val.strip() else default


def clean_int(val: Any, default: int = 10) -> int:
    """Helper to ensure clean int when called directly or via Query."""
    if isinstance(val, int):
        return val
    try:
        return int(val)
    except Exception:
        return default


def parse_list_param(val: Any) -> Optional[List[str]]:
    """Helper to parse comma-separated query parameters."""
    if not val or not isinstance(val, str):
        return None
    items = [x.strip() for x in val.split(",") if x.strip()]
    return items if items else None


# --------------------------------- Frontend ----------------------------------
@app.get("/", response_class=HTMLResponse, tags=["Frontend"])
def serve_dashboard():
    """Serves the Single Page Application dashboard frontend."""
    if INDEX_HTML_PATH.exists():
        return HTMLResponse(content=INDEX_HTML_PATH.read_text(encoding="utf-8"))
    return HTMLResponse(
        content="""
        <html>
            <body style='font-family:sans-serif;padding:40px;text-align:center;'>
                <h2>Sales Analytics API Backend is Running!</h2>
                <p>To view documentation, visit <a href="/docs">/docs</a>.</p>
            </body>
        </html>
        """
    )


# -------------------------------- Health & Info ------------------------------
@app.get("/api/health", tags=["System"])
def get_health():
    """Health check endpoint showing dataset status and load timestamp."""
    return {
        "status": "online",
        "dataset_file": db.filename,
        "total_records": len(db.df),
        "last_loaded": db.last_loaded,
        "uptime": round(time.time() - db.last_loaded, 2),
    }


# ------------------------------ Filters Metadata -----------------------------
@app.get("/api/filters", tags=["Metadata"])
def get_filter_options():
    """Returns dynamic date bounds, regions, categories, and product lists."""
    return db.get_filter_options()


# -------------------------------- KPIs Endpoint ------------------------------
@app.get("/api/kpis", tags=["Analytics"])
def get_kpis(
    start_date: Optional[str] = Query(None, description="Start date YYYY-MM-DD"),
    end_date: Optional[str] = Query(None, description="End date YYYY-MM-DD"),
    regions: Optional[str] = Query(None, description="Comma-separated regions"),
    categories: Optional[str] = Query(None, description="Comma-separated categories"),
    product: Optional[str] = Query(None, description="Exact product name"),
):
    """Computes Revenue, Orders, Units, AOV, and Discount KPIs on filtered data."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_kpis(fdf)


# -------------------------------- Trend Chart --------------------------------
@app.get("/api/charts/trend", tags=["Charts"])
def get_trend_chart(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
):
    """Returns monthly time series of revenue, orders, and units."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_trend(fdf)


# ------------------------------ Category Chart -------------------------------
@app.get("/api/charts/category", tags=["Charts"])
def get_category_chart(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
):
    """Returns category breakdown metrics and revenue share percentages."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_category_distribution(fdf)


# ------------------------------ Regional Chart -------------------------------
@app.get("/api/charts/regional", tags=["Charts"])
def get_regional_chart(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
):
    """Returns regional performance cross-tabulation across categories."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_regional_breakdown(fdf)


# ------------------------------ Weekday Chart --------------------------------
@app.get("/api/charts/weekday", tags=["Charts"])
def get_weekday_chart(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
):
    """Returns order and sales distribution across days of the week."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_weekday_dynamics(fdf)


# ---------------------------- Top Products Chart -----------------------------
@app.get("/api/charts/top-products", tags=["Charts"])
def get_top_products_chart(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
    top_n: int = Query(10, ge=1, le=50),
    metric: str = Query("revenue", pattern="^(revenue|units)$"),
):
    """Returns top N products sorted by revenue or quantity."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_top_products(fdf, limit=clean_int(top_n, 10), metric=clean_str(metric, "revenue"))


# --------------------------- Top Customers Chart -----------------------------
@app.get("/api/charts/top-customers", tags=["Charts"])
def get_top_customers_chart(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
    top_n: int = Query(10, ge=1, le=50),
):
    """Returns top N customer accounts sorted by total spend."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_top_customers(fdf, limit=clean_int(top_n, 10))


# -------------------------------- Insights -----------------------------------
@app.get("/api/charts/insights", tags=["Analytics"])
def get_insights(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
):
    """Returns executive superlatives and key drivers."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_executive_insights(fdf)


# ----------------------------- Orders Explorer -------------------------------
@app.get("/api/orders", tags=["Orders"])
def get_orders(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
    search: Optional[str] = None,
    sort_by: str = Query("Order_Date"),
    sort_dir: str = Query("desc", pattern="^(asc|desc)$"),
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
):
    """Searchable, sortable, paginated orders ledger."""
    fdf = db.filter_df(
        start_date=clean_str(start_date),
        end_date=clean_str(end_date),
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=clean_str(product),
    )
    return db.get_paginated_orders(
        fdf,
        search=clean_str(search),
        sort_by=clean_str(sort_by, "Order_Date"),
        sort_dir=clean_str(sort_dir, "desc"),
        page=clean_int(page, 1),
        page_size=clean_int(page_size, 20),
    )


# -------------------------------- CSV Export ---------------------------------
@app.get("/api/export", tags=["Export"])
def export_csv(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    regions: Optional[str] = None,
    categories: Optional[str] = None,
    product: Optional[str] = None,
):
    """Exports current filtered records to downloadable CSV."""
    fdf = db.filter_df(
        start_date=start_date,
        end_date=end_date,
        regions=parse_list_param(regions),
        categories=parse_list_param(categories),
        product=product,
    )
    buffer = db.export_csv(fdf)
    return StreamingResponse(
        buffer,
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=filtered_sales_data.csv"},
    )


# ------------------------------- Data Upload ---------------------------------
@app.post("/api/upload", tags=["Dataset Management"])
async def upload_dataset(file: UploadFile = File(...)):
    """Upload a new Excel (.xlsx) or CSV file to dynamically reload the dashboard."""
    if not file.filename.lower().endswith((".xlsx", ".xls", ".csv")):
        raise HTTPException(
            status_code=400,
            detail="Unsupported file format. Please upload an Excel (.xlsx/.xls) or CSV (.csv) file.",
        )

    try:
        content = await file.read()
        stream = io.BytesIO(content)
        db.load_dataset(source=stream, filename=file.filename)
        return {
            "status": "success",
            "message": f"Dataset '{file.filename}' successfully loaded with {len(db.df):,} records.",
            "total_records": len(db.df),
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# =============================================================================
# 4. RUNNER & AUTOMATION
# =============================================================================
import socket


def is_port_in_use(port: int, host: str = "127.0.0.1") -> bool:
    """Check if a port is currently occupied."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex((host, port)) == 0


def find_available_port(start_port: int = 8000, host: str = "127.0.0.1", max_attempts: int = 20) -> int:
    """Find the first available port starting from start_port."""
    for p in range(start_port, start_port + max_attempts):
        if not is_port_in_use(p, host):
            return p
    return start_port


def open_browser_delayed(url: str, delay: float = 1.2):
    """Open default web browser after server starts."""
    time.sleep(delay)
    try:
        webbrowser.open(url)
    except Exception as e:
        print(f"[Warning] Could not open browser automatically: {e}")


def start_server(host: str = "127.0.0.1", port: int = 8000):
    """Start the Uvicorn web server hosting FastAPI and Frontend with auto port fallback."""
    actual_port = find_available_port(port, host)
    if actual_port != port:
        print(f"[Notice] Port {port} is currently busy. Automatically switched to available port {actual_port}.")
    port = actual_port

    url = f"http://{host}:{port}"
    print("=" * 70)
    print("[SalesPulse Enterprise Dashboard - Full-Stack Application]")
    print(f"[*] Local Web App : {url}")
    print(f"[*] Interactive API Docs (Swagger): {url}/docs")
    print(f"[*] Dataset Status: {len(db.df):,} records loaded from '{db.filename}'")
    print("=" * 70)

    # Launch browser automatically in background thread
    threading.Thread(target=open_browser_delayed, args=(url,), daemon=True).start()

    # Run Uvicorn
    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    start_server()
