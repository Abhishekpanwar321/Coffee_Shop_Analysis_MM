# ☕ Coffee Orders Sales Dashboard

> An end-to-end Excel data analytics project transforming raw coffee order data into an interactive business intelligence dashboard — revealing sales trends, top customers, and country-level performance across 3 markets.

---

## 📋 Description

This project analyses **1,000 real-world coffee orders** spanning January 2019 to August 2022 across the United States, Ireland, and the United Kingdom. Starting from three raw, disconnected spreadsheet tables (orders, customers, products), the project builds a fully automated, pivot-driven Excel dashboard that gives coffee retail managers instant visibility into revenue performance, product mix, and customer behaviour.

The project demonstrates a complete data analytics workflow — from messy raw data to a polished, filterable dashboard — using only Microsoft Excel.

---

## 🧩 Problem Solved

Coffee retail managers were working with fragmented spreadsheet data: orders, customers, and products lived in separate sheets with no relationships, no aggregations, and no visualisations. Answering basic questions like *"Which coffee type drove the most revenue last quarter?"* or *"Who are our top 5 customers?"* required manual, error-prone calculations each time.

**This project eliminates that friction** by:
- Joining all three data tables using XLOOKUP / INDEX-MATCH formulas
- Auto-calculating revenue (Quantity × Unit Price) across all 1,000 rows
- Surfacing insights through interactive Pivot Tables and Charts with slicers

---

## 📈 Impact

| Metric | Result |
|---|---|
| **Total Revenue Analysed** | $45,134 across 3.5 years |
| **Orders Processed** | 1,000 orders, 913 unique customers |
| **Markets Covered** | United States · Ireland · United Kingdom |
| **Top Market** | 🇺🇸 United States — $35,639 (79% of revenue) |
| **Best-Selling Coffee** | Excelsa — $12,306 total sales |
| **Top Customer** | Allis Wilmore — $317.07 lifetime spend |
| **Time to Insight** | From raw data → dashboard in one workbook |

---

## ✨ Features

- **Automated Data Join** — XLOOKUP formulas link Customer Name, Email, Country, Coffee Type, Roast Type, Size, Unit Price, and Loyalty Card status directly from source tables into the orders sheet — zero manual copying
- **Revenue Calculation** — Derived `Sales` column (Quantity × Unit Price) calculated across all 1,000 rows
- **Total Sales Over Time Pivot** — Monthly revenue broken down by all 4 coffee types (Arabica, Excelsa, Liberica, Robusta) across 2019–2022
- **Country Bar Chart** — Visual comparison of revenue across US, Ireland, and UK
- **Top 5 Customers Chart** — Ranked customer leaderboard by lifetime spend
- **Interactive Slicers** — Filter by Roast Type, Package Size, and Loyalty Card status with one click
- **Timeline Slicer** — Drill into any date range dynamically
- **Named Full Coffee Types** — Human-readable names (e.g. "Arabica" instead of "Ara") via lookup columns
- **Dashboard Sheet** — All charts and slicers assembled on a single, clean view

---

## 🛠️ Technology Stack

| Tool | Purpose |
|---|---|
| **Microsoft Excel** | Primary analytics and dashboard platform |
| **Excel Pivot Tables** | Data aggregation and cross-tabulation |
| **Excel Pivot Charts** | Line chart (sales over time), bar chart (country), bar chart (top customers) |
| **XLOOKUP** | Joining customer and product data into the orders table |
| **IF / IFS** | Expanding coded coffee/roast type abbreviations to full names |
| **Timeline Slicer** | Date-range filtering on Pivot Tables |
| **Pivot Slicers** | Categorical filtering (roast type, size, loyalty card) |

---

## 🚀 Setup / Installation

**Requirements:** Microsoft Excel 2019+ or Microsoft 365 (XLOOKUP requires Excel 2019+)

### Option A — Use the completed project file
1. Download **`coffeeOrdersProject.xlsx`**
2. Open in Excel — all formulas, pivot tables, and the dashboard are pre-built
3. Navigate to the **Dashboard** sheet to explore

### Option B — Rebuild from raw data (follow the development steps)
1. Download **`coffeeOrdersData.xlsx`** (raw source data)
2. Follow the **Development Steps** section below to recreate the project from scratch
3. This is the recommended approach for learning

---

## 📖 Usage

### Navigating the Dashboard
| Sheet | Contents |
|---|---|
| `orders` | Full enriched orders table with all XLOOKUP-joined fields |
| `customers` | Source customer records (1,000 customers) |
| `products` | Source product catalogue (48 SKUs across 4 coffee types) |
| `TotalSales` | Pivot: monthly sales by coffee type (2019–2022) |
| `CountryBarChart` | Pivot: total revenue by country |
| `Top5Customers` | Pivot: top 5 customers by spend |
| `Dashboard` | All charts + slicers assembled for end-user use |

### Using the Slicers
On the Dashboard sheet, use the slicers to interactively filter all charts simultaneously:
- **Roast Type** — Filter by Dark, Light, or Medium roast
- **Size** — Filter by package size (0.2kg, 0.5kg, 1.0kg, 2.5kg)
- **Loyalty Card** — Compare card holders vs non-holders
- **Timeline** — Select any custom date range from Jan 2019 – Aug 2022

---

## 🔄 Development Steps

### Phase 1 — Understanding the Business Problem
- Identified that raw data was split across 3 disconnected sheets
- Defined the key business questions: revenue trends, top markets, best customers, product mix
- Planned the target output: a single-sheet interactive dashboard

### Phase 2 — Data Preparation & Joining (orders sheet)
- Audited the raw `orders` sheet — identified 8 columns needing data from other tables
- Used **XLOOKUP** to pull `Customer Name`, `Email`, `Country`, `Loyalty Card` from `customers`
- Used **XLOOKUP** to pull `Coffee Type`, `Roast Type`, `Size`, `Unit Price` from `products`
- Added `Sales` column: `=Quantity × Unit Price`
- Added `Coffee Type Name` and `Roast Type Name` using **IFS** to expand abbreviations to full names
- Validated all 1,000 rows for completeness and correctness

### Phase 3 — Building Pivot Tables
- Created **TotalSales** pivot: rows = Year/Month, columns = Coffee Type Name, values = Sum of Sales
- Created **CountryBarChart** pivot: rows = Country, values = Sum of Sales
- Created **Top5Customers** pivot: rows = Customer Name, values = Sum of Sales, top-5 filter applied

### Phase 4 — Building Charts
- Line/Area chart on TotalSales pivot — sales over time by coffee type
- Horizontal bar chart on CountryBarChart pivot — revenue by market
- Horizontal bar chart on Top5Customers pivot — ranked customer spend

### Phase 5 — Adding Interactivity
- Inserted **Timeline Slicer** connected to all pivot tables
- Inserted **Slicers** for Roast Type, Size, and Loyalty Card
- Connected all slicers to all three pivot tables for synchronised filtering

### Phase 6 — Dashboard Assembly
- Created a dedicated `Dashboard` sheet
- Moved all charts and slicers onto the Dashboard sheet
- Applied consistent formatting, titles, and a clean layout

### Phase 7 — Testing & Refinement
- Tested all slicer combinations for correct chart updates
- Verified revenue totals against manual spot-checks
- Confirmed top 5 customers and country totals match expected values

---

## 🤝 Contribution Guidelines

Contributions are welcome! To suggest improvements:
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit your changes: `git commit -m 'Add: your feature description'`
4. Open a Pull Request with a clear description of the change

Ideas for future enhancements:
- Add a profit margin analysis using the `Profit` column in `products`
- Build a customer segmentation view by loyalty card status
- Create a Year-over-Year growth chart
- Migrate dashboard to Power BI for richer interactivity

---

## 📄 License

This project is licensed under the **MIT License** — you are free to use, copy, modify, and distribute this project with attribution.

---

*Built with Microsoft Excel · Data spans January 2019 – August 2022 · 1,000 orders · 3 markets*
