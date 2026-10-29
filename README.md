# SalesPulse Enterprise - Full-Stack Sales Analytics Platform

A production-grade, full-stack Business Intelligence & Sales Analytics Dashboard built in Python with a dedicated **Backend REST API** and a modern, reactive **Frontend SPA**.

---

## 🌟 Architecture Overview

This project implements a clean separation of concerns between backend data processing and frontend presentation:

```
Course/
├── app.py                # Full-Stack Backend (FastAPI, Uvicorn, Pandas vectorized data layer)
├── chart.py              # Universal runner (python chart.py OR streamlit run chart.py)
├── requirements.txt      # Dependency specifications
├── sales.xlsx            # Primary dataset (50,000 transaction records)
├── templates/
│   └── index.html        # Modern, responsive Frontend SPA (Tailwind CSS, Chart.js, Lucide)
└── README.md             # Documentation and usage guide
```

### 1. Backend Layer (`app.py`)
- **FastAPI & Uvicorn**: High-performance asynchronous REST API server.
- **Pandas Vectorized Engine**: In-memory caching and lightning-fast (<15ms) data querying across 50,000+ orders.
- **RESTful Endpoints**:
  - `GET /` — Serves the interactive Single Page Application frontend.
  - `GET /api/health` — System health check, uptime, dataset statistics.
  - `GET /api/filters` — Dynamic filter options (date range, regions, categories, products).
  - `GET /api/kpis` — Key business metrics (Total Revenue, Orders, Units, AOV, Avg Discount, Discount Savings).
  - `GET /api/charts/trend` — Monthly & yearly time-series revenue and order volume trends.
  - `GET /api/charts/category` — Category distribution and revenue share percentages.
  - `GET /api/charts/regional` — Region × Category revenue cross-tabulation matrix.
  - `GET /api/charts/weekday` — Day-of-the-week ordering traffic and volume analysis.
  - `GET /api/charts/top-products` — Top N product rankings by revenue or units sold.
  - `GET /api/charts/top-customers` — Top VIP customer accounts by total expenditure.
  - `GET /api/charts/insights` — Algorithmic business intelligence (top region, star product, peak month, largest single transaction).
  - `GET /api/orders` — Searchable, sortable, paginated order ledger with instant debouncing.
  - `GET /api/export` — Streamed CSV export of currently filtered data.
  - `POST /api/upload` — Dynamic Excel (.xlsx) or CSV upload to reload the dataset in real time without server restart.
- **Auto-generated Swagger Docs**: Available at `http://127.0.0.1:8000/docs`.

### 2. Frontend Layer (`templates/index.html`)
- **Modern Responsive Design**: Built with Tailwind CSS and glassmorphism styling.
- **Interactive Visualizations**: 6 high-resolution Chart.js charts with custom tooltips, gradients, and animated transitions.
- **Theme Support**: Seamless toggle between Dark Mode and Light Mode.
- **Real-Time Client State**:
  - Instant multi-filter reactivity (Date Range, Quick Presets: 2026, 2025, 2024, 2023, Last 90 Days).
  - Region & Category pill toggles with multi-select support and "Select All".
  - Dynamic cascading Product dropdown.
  - Top N slider (5, 10, 15, 20 items).
  - Table live search across Customer Name, Order ID, and Product.
  - Column header sorting (Ascending / Descending) and pagination (10, 20, 50, 100 rows per page).
  - One-click CSV export button.
  - File upload modal with drag-and-drop zone.

---

## 🚀 How to Run the Dashboard

### Prerequisites
Make sure your dependencies are installed:
```bash
pip install -r requirements.txt
```

### Option A: Launch Full-Stack Application (Recommended)
Run:
```bash
python app.py
```
*Or interchangeably:*
```bash
python chart.py
```

**What happens:**
1. The backend initializes and loads `sales.xlsx` into high-speed memory.
2. The Uvicorn server starts on `http://127.0.0.1:8000`.
3. Your default web browser opens automatically to the dashboard!
4. You can explore the interactive API documentation at `http://127.0.0.1:8000/docs`.

---

### Option B: Run Streamlit Interface
If you want to view the Streamlit interface:
```bash
streamlit run chart.py
```

---

## 📊 Business Metrics Computed
| Metric | Calculation | Formatted Display |
| :--- | :--- | :--- |
| **Total Revenue** | $\sum \text{Total\_Amount\_INR}$ | ₹ Cr / ₹ L / ₹ K |
| **Total Orders** | $\text{Distinct Count}(\text{Order\_ID})$ | Formatted integer |
| **Units Sold** | $\sum \text{Quantity}$ | Total items |
| **Avg Order Value (AOV)** | $\frac{\text{Total Revenue}}{\text{Total Orders}}$ | ₹ Amount per order |
| **Avg Discount** | $\text{Mean}(\text{Discount\_Percent})$ | Percentage % |
| **Discount Savings** | $\sum (\text{Qty} \times \text{Unit\_Price} - \text{Total})$ | ₹ Value saved |
