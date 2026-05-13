#  Lounge Financial Reporting System

This document outlines the architecture, data flow, and business logic of the Lounge Report Financial Engine and its associated interactive dashboard.

## 1. Project Overview
The Lounge Financial Reporting System is an end-to-end analytical tool designed to process raw transaction volume data for services (Domestic, Railway, and Global lounges), apply complex pricing matrices to calculate revenue and costs, and present the financial margins in a modern, interactive web dashboard.

### Core Objectives:
- Map millions of transactions to dynamic, partner-specific pricing agreements.
- Compute accurate Revenue, Cost, and Margin data.
- Automatically identify and flag anomalous transactions (e.g., expired contracts, missing rates).
- Provide a responsive, executive-level visualization dashboard.

---

## 2. System Architecture

The project consists of two primary components: a Python-based backend data processor and a Vanilla HTML/JS frontend dashboard.

### 2.1 Backend: The Financial Engine (`processor.py`)
The Python script acts as the financial engine. It reads raw CSV data, cross-references it with master pricing spreadsheets, and outputs a structured JSON file.

**Inputs:**
- `Data.csv`: The raw transaction volume data. Key columns include `transaction_recorded_time_IST`, `outlet_id_long`, `partner_name`, `service_name`, `free_count`, and `Paid Visit`.
- `Domestic Lounge.xlsx`: Contains partner-specific sheets and a global `CostSheet` for Domestic lounges.
- `Railway Lounge.xlsx`: Contains partner-specific sheets and a global `CostSheet` for Railway lounges.
- `Global Lounge.xlsx`: A complex master file containing `Region Criteria`, `Sales Rate`, `Cost Sheet`, and specialized sheets for partners like `ETT`, `LoungePair`, and `Unimoni`.

**Core Processing Logic:**
1. **Data Normalization:** Parses dates, handles partner name standardizations (e.g., Vodafone), and excludes specific known-bad outlet/partner combinations.
2. **Dynamic Rate Resolution:** The system handles *historical pricing*. If an outlet has multiple rates over time, the engine searches all rates for that `Outlet ID` and dynamically selects the correct price based on the transaction date (`txn_date <= expiry`).
3. **Financial Math:**
   - `Total PAX` = `free_count` + `Paid Visit` (with partner-specific exceptions like IDFC Bank).
   - `Line Revenue` = `Relevant Rate` × `Total PAX`.
   - `Line Cost` = `Relevant Cost` × `Total PAX`.
   - `Margin` = `Line Revenue` - `Line Cost`.
4. **Debug Flagging:** Transactions without valid rates are flagged with specific `debug_reason` codes (e.g., `Rate expired or not yet valid`, `Cost missing`, `Unknown partner`).

**Output:**
- `final_data.json`: The fully processed analytical payload containing KPIs, monthly aggregations, top 10 outlet lists, debug records, and slimmed-down transaction rows.

---

### 2.2 Frontend: The Dashboard (`Lounge_Report_Final_V10.html`)
The frontend is a standalone, dynamic HTML file that consumes `final_data.json`. 

**Key Features:**
- **Modern Dark UI:** A premium, glassmorphism-inspired dark theme for executive presentations.
- **Cross-Filtering:** Users can filter data by `Service Type` (Domestic, Global, Railway) and `Partner`.
- **Top Outlet Analytics:** Interactive data tables displaying the top-performing outlets sorted by Margin or PAX.
- **Debug & Action Log:** A dedicated "Changes/Debug" view that isolates transactions flagged by the backend (e.g., expired contracts) so operational teams can take immediate action to update the master pricing sheets.

#### Dashboard Charts & Visualizations
The V10 dashboard uses the **ApexCharts** library to render dynamic, interactive visualizations. These charts are split across different tabs to serve specific analytical purposes:

**Overview Tab (Volume & Growth Analytics):**
1. **PAX Trend (Area Chart):** Shows the overall passenger volume (PAX) trend across all available months. Used to track long-term business growth.
2. **Daily PAX Comparison (Overlaid Area Chart):** Compares the daily passenger volume of a selected month against the previous month. Used to identify seasonal spikes or sudden drops in footfall.

**Financial Tab (Revenue & Margin Analytics):**
3. **Monthly Revenue vs Cost (Donut Charts):** Displays up to 3 separate donut charts for the most recent months. Shows the proportion of Cost vs. Margin, with the absolute Margin % prominently displayed in the center. Used for high-level profitability checks.
4. **Top 10 Outlets by Margin (Horizontal Bar Chart):** Ranks the top 10 most profitable outlets. Handles massive outliers gracefully to ensure smaller outlets remain visible. Used to identify key revenue drivers.
5. **Daily Revenue vs Cost Breakout (Line Chart):** Plots daily Revenue and Cost side-by-side for a selected month. Used to track daily financial performance and ensure costs do not outpace revenue on a day-to-day basis.
*(Note: Financial charts automatically split into separate INR (₹) and USD ($) variants depending on the applied filters).*

**Outlets Tab (Deep-Dive Analytics):**
6. **Outlet-Specific Daily Breakout (Line Chart):** When a specific outlet is clicked in the data table, this chart dynamically renders that single outlet's daily revenue, cost, and margin over the selected month. Used for granular, outlet-level troubleshooting and performance tracking.

---

## 3. Data Dictionary & Debug Codes

### Common Debug Reason Codes
When the engine cannot calculate a clean margin, it assigns one of the following codes:
- **`Rate expired or not yet valid` / `Cost contract expired`**: The transaction occurred after the `Expiry Date` or `Valid Till` date in the master sheet.
- **`No rate found for partner`**: The partner exists, but the specific `Outlet ID` is missing from their pricing sheet.
- **`Unknown partner — no master rate`**: The `partner_name` in `Data.csv` does not match any known sheet in the master Excel files.
- **`Cost missing — outlet not in cost sheet`**: The outlet generated revenue but has no mapped cost in the `CostSheet`.
- **`Outlet excluded — non-operational`**: The outlet is flagged as not "Live" in the Global Cost sheet.

### Currency Handling
- **Domestic & Railway:** All calculations and aggregations are strictly processed in `INR (₹)`.
- **Global:** All calculations are strictly processed and reported in `USD ($)`.

---

## 4. Operational Workflow (How to Update)

1. **Update Master Data:** When contracts are renewed, add the new rows to the respective Excel file (`Global Lounge.xlsx`, etc.). Ensure the new row has the updated `Price` and the new `Expiry Date`. Do not delete the historical rows; the system will automatically handle date overlaps.
2. **Refresh Raw Data:** Replace `Data.csv` with the latest transaction dump.
3. **Run Engine:** Execute `python processor.py`. This will generate an updated `final_data.json`.
4. **View Dashboard:** Open `Lounge_Report_Final_V10.html` in any modern web browser to view the updated financial analytics.
