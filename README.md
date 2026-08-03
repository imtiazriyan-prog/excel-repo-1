# 🛒 Superstore Sales Dashboard — Excel Analytics Project

> **A complete end-to-end data analysis of 4 years of US retail sales data (2015–2018), built entirely in Microsoft Excel with interactive dashboards, pivot-driven charts, XLOOKUP lookups, and KPI summaries.**

---

## 📌 Project Overview

This project analyzes the **Sample Superstore** dataset — a widely used retail dataset covering orders across the United States. The goal was to extract actionable business insights on **sales performance, shipping behavior, regional trends, and customer segmentation** using Excel as the sole analytics tool.

| Metric | Value |
|--------|-------|
| 📅 Date Range | Jan 2015 – Dec 2018 |
| 📦 Total Orders | 4,922 unique orders |
| 👥 Unique Customers | 793 |
| 💰 Total Revenue | $2,261,536 |
| 🏆 Top Category | Technology |
| 🗺 Top Region | West |
| 🗂 Product Sub-Categories | 17 |

---

## 📂 Repository Structure

```
superstore-excel-dashboard/
│
├── 📊 Superstore_Dashboard_Final.xlsx    # Main Excel workbook (all analysis + dashboard)
│
├── 📁 data/
│   └── superstore_dataset_info.md        # Dataset schema and field descriptions
│
├── 📁 analysis/
│   └── key_findings.md                   # Written summary of insights and recommendations
│
├── 📁 assets/
│   └── dashboard_preview.png             # Screenshot of the final dashboard
│
├── 📁 docs/
│   └── methodology.md                    # Approach, formulas used, and design decisions
│
└── README.md                             # This file
```

---

## 📊 Excel Workbook — Sheet Breakdown

The workbook (`Superstore_Dashboard_Final.xlsx`) contains **11 sheets**, each serving a specific analytical purpose:

| Sheet | Type | Purpose |
|-------|------|---------|
| `Dataset` | Raw Data | Source data with 9,800 rows and 18 core fields + computed KPI columns |
| `XLOOKUP` | Formula Sheet | Dynamic product/order lookup using Excel's XLOOKUP function |
| `Sales Trend` | Pivot Table | Month-over-month revenue trend across all regions (2015–2018) |
| `Sales by Region` | Pivot Table | Revenue breakdown by US region (East, West, Central, South) |
| `Sales by Category` | Pivot Table | Revenue split across Furniture, Office Supplies, Technology |
| `Sub Category` | Pivot Table | Revenue by all 17 product sub-categories |
| `Customer Type` | Pivot Table | Segment analysis — Consumer, Corporate, Home Office |
| `Shipping Behaviour` | Pivot Table | Orders and sales by Ship Mode |
| `Bar Chart` | Chart Data | Aggregated data feeding the sales bar chart visual |
| `Dashboard` | Dashboard | Final interactive dashboard with slicers and charts |
| `Dashboard Layout` | Planning | Layout blueprint used to design the dashboard |

---

## 🔑 Key Insights

### 💰 Sales Performance
- **Total revenue: $2.26M** across 4 years with consistent YoY growth.
- **Technology** is the highest-grossing category, followed by Furniture and Office Supplies.
- The **West region** leads in total sales; the **South** has the most room for growth.

### 🚚 Shipping Behaviour
- **Standard Class** is the most used shipping method (~60% of orders).
- **Same Day** shipping accounts for the smallest share but typically carries higher-value orders.
- Average shipping time varies meaningfully by Ship Mode — an opportunity to align customer expectations.

### 👤 Customer Segments
- **Consumer** segment drives the majority of orders.
- **Corporate** customers generate higher average order values.
- **Home Office** is the smallest but growing segment.

### 📦 Product Categories
- **Chairs and Phones** are top sub-categories by revenue.
- **Fasteners and Labels** have the lowest revenue but are high-frequency items.
- Office Supplies show the highest order volume but lowest average sale value.

---

## 🛠 Excel Features & Techniques Used

| Feature | Where Used |
|---------|-----------|
| **Pivot Tables** | Sales Trend, Region, Category, Sub-Category, Customer Type, Shipping |
| **XLOOKUP** | Dynamic order-level lookups across the dataset |
| **Pivot Charts** | Bar chart, line trend chart, pie/donut charts |
| **Slicers** | Region, Segment, Category filters on dashboard |
| **Conditional Formatting** | KPI highlighting and data bars |
| **Named Ranges** | Used for clean formula references |
| **YYYY-MM Column** | Custom date field for time-series grouping |
| **KPI Summary Cells** | Total Sales, Total Orders, Average Sales, Top Category, Top Region |
| **Dashboard Layout** | Dedicated planning sheet used before building the final view |

---

## 📐 Dataset Schema

The `Dataset` sheet contains the following fields:

| Column | Type | Description |
|--------|------|-------------|
| Row ID | Integer | Unique row identifier |
| Order ID | String | Unique order identifier |
| Order Date | Date | Date the order was placed |
| Ship Date | Date | Date the order was shipped |
| Ship Mode | String | Shipping class (Standard, Second, First, Same Day) |
| Customer ID | String | Unique customer identifier |
| Customer Name | String | Full name of the customer |
| Segment | String | Customer type (Consumer / Corporate / Home Office) |
| Country | String | Always "United States" |
| City | String | City of delivery |
| State | String | US state of delivery |
| Postal Code | Integer | ZIP code |
| Region | String | US region (East / West / Central / South) |
| Product ID | String | Unique product identifier |
| Category | String | Product category (Furniture / Office Supplies / Technology) |
| Sub-Category | String | Product sub-category (17 types) |
| Product Name | String | Full product name |
| Sales | Float | Revenue value for that line item |
| YYYY-MM | String | Custom date field for time-based grouping |

**Computed KPI Columns (added during analysis):**

| Column | Formula Used | Description |
|--------|-------------|-------------|
| Total Sales | `=SUM(Sales column)` | Aggregate revenue |
| Total Order | `=COUNTA(Order ID)` | Count of all orders |
| Average Sales | `=AVERAGE(Sales column)` | Mean order value |
| Top Category | `=INDEX/MATCH` or manual | Best-performing category |
| Top Region | `=INDEX/MATCH` or manual | Best-performing region |

---

## 🚀 How to Use This File

### Prerequisites
- Microsoft Excel 2019 or later (for XLOOKUP support)
- Excel 365 recommended for full slicer and dynamic array support

### Steps
1. **Clone or download** this repository.
2. Open `Superstore_Dashboard_Final.xlsx` in Excel.
3. Navigate to the **`Dashboard`** sheet to view the interactive summary.
4. Use the **slicers** (Region, Segment, Category) to filter all charts simultaneously.
5. Explore individual pivot sheets for deeper drill-down.
6. The **`XLOOKUP`** sheet lets you look up any order by entering an Order ID or Product ID.

> ⚠️ **Note:** If pivot tables show stale data, right-click any pivot → **Refresh All** to update.

---

## 📈 Dashboard Preview

> *(Add a screenshot here after exporting the dashboard view)*

```
[ Insert dashboard_preview.png here ]
```

To capture: Go to `Dashboard` sheet → `File > Export > Save as PDF` or take a screenshot and save as `assets/dashboard_preview.png`.

---

## 💡 Future Improvements

- [ ] Add **Profit and Discount** columns to the dataset for margin analysis
- [ ] Build a **customer lifetime value (CLV)** calculation
- [ ] Add **year-over-year growth %** as a calculated KPI
- [ ] Create a **dynamic top-N products** view using LARGE() and INDEX/MATCH
- [ ] Migrate dashboard to **Power BI** for web sharing and real-time refresh
- [ ] Add **state-level map visualization** using Excel's map chart type

---

## 🗂 Related Resources

- [Sample Superstore Dataset (Tableau Public)](https://public.tableau.com/app/learn/sample-data)
- [Microsoft Excel XLOOKUP Documentation](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)
- [Excel Pivot Table Guide](https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576)

---

## 👤 Author

Imtiaz Ahmad
Data analyst who turns raw data into clear insights using a strong toolkit that includes Microsoft Excel, SQL, Python, Microsoft Power BI, and Tableau. Combine technical skills with analytical thinking to build dashboards, automate reporting, uncover trends, and support smarter business decisions.
- 📧 Email: imtiazahmadriyan@gmail.com
- 💼 LinkedIn: https://www.linkedin.com/in/imtiaz-ahmad-a716a15a
- 🐙 GitHub: https://github.com/imtiazriyan-prog

---

