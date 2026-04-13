# 📊 Sales MIS Dashboard — Excel

An end-to-end MIS reporting project built in Microsoft Excel on real-world style sales transaction data. Covers the complete MIS workflow — raw data processing, formula logic, pivot table analysis, and an interactive dark-theme sales dashboard.

> **Tools Used:** Microsoft Excel
> **Skills Demonstrated:** Data Cleaning, VLOOKUP, IF/Nested IF, Pivot Tables, Calculated Fields, Dashboard Design

---

## 📁 File
`Sales_MIS_Dashboard_Excel.xlsx`

---

## 📌 Project Overview

This project simulates a real MIS reporting task where raw invoice data from a printer products company is cleaned, enriched with lookup data, analyzed using pivot tables, and presented as an interactive executive dashboard.

The dataset includes:
- **Sales invoices and return transactions** across 6 US states
- **19 customers** across Automotive and Financial industries
- **Products** across 3 lines — Accessories, Consumables, and Printers

---

## 🗂️ Sheet Structure

| Sheet | Description |
|-------|-------------|
| `Invoices` | Main transaction data with all calculated columns |
| `Products` | Product master data (Part#, Description, Product Line, Cost) |
| `Customers` | Customer master data (Account, Name, State, Type) |
| `Region vs Cust Type` | Pivot — Revenue by Region and Customer Type |
| `Region vs Product Line` | Pivot — Revenue by Region and Product Line |
| `Product Analysis` | Pivot — Cost vs Revenue by Region and Customer Type |
| `Profit_Margin` | Overall Profit Margin calculation |
| `Return_Analysis Sheet` | Pivot + Chart — Returns broken down by Reason Code |
| `DashBoard` | Interactive Sales Dashboard with KPIs, Charts, Slicers and Timeline |

---

## ✅ Tasks Completed

| Sr. No. | Task | Column | Tool/Method |
|---------|------|--------|-------------|
| 1 | Negative value for return transaction | New Unit Price | IF |
| 2 | Categorise product — Premium, Basic, Free, Return | Product Value | Nested IF |
| 3 | Get Unit Cost from Products table | Unit Cost | VLOOKUP |
| 4 | Get Customer Type from Customers table | Customer Type | VLOOKUP |
| 5 | Get Customer Region from Customers table | Customer Region | VLOOKUP |
| 6 | Get Product Line from Products table | Product Line | VLOOKUP |
| 7 | Calculate Cost | Cost | Formula |
| 8 | Calculate Revenue | Revenue | Formula |
| 9 | Calculate Profit | Profit | Formula |
| 10 | Calculate Profit Margin | Profit Margin | Calculated Field |
| 11 | Region vs Customer Type analysis | — | Pivot Table |
| 12 | Region vs Product Line analysis | — | Pivot Table |
| 13 | Free Product Analysis | — | Pivot Table |
| 14 | Value of Product Analysis | — | Pivot Table |
| 15 | Extract month from date | Month | TEXT Formula |
| 16 | Extract quarter from date | Quarter | ROUNDUP & MONTH Formula |
| 17 | Calculate profit margin per row | Profit Margin % | Formula |
| 18 | KPI — Total Revenue | Dashboard | SUM Formula |
| 19 | KPI — Total Cost | Dashboard | SUM Formula |
| 20 | KPI — Total Profit | Dashboard | SUM Formula |
| 21 | KPI — Overall Profit Margin % | Dashboard | SUM Formula |
| 22 | Return analysis by reason code | Return Analysis Sheet | Pivot Table |
| 23 | Returns by reason code visual | Return Analysis Sheet | Pivot Chart |
| 24 | Interactive date filter | Dashboard | Timeline Slicer |
| 25 | Interactive filters for all charts | Dashboard | Slicers |
| 26 | Sales Dashboard with dark theme | Dashboard Sheet | Charts + Formatting |

---

## 📊 Dashboard Features

- **KPI Cards** — Total Revenue, Total Cost, Total Profit, Overall Profit Margin %
- **3 Interactive Charts** — Region vs Customer Type, Region vs Product Line, Product Analysis (Cost vs Revenue)
- **Return Analysis Chart** — Returns broken down by Reason Code
- **Timeline Slicer** — Filter all charts by date range
- **4 Slicers** — Customer Type, Customer Region, Product Value, Product Line
- **Dark Theme** — Professional dark navy blue dashboard design

---

## 🔑 Key Business Insights

- **Total Revenue:** $9,32,948 | **Total Profit:** $4,89,495 | **Profit Margin:** 52%
- **Georgia** is the highest revenue generating region at $4,19,600
- **Products** contribute the highest revenue across all regions compared to Accessories and Consumables
- **Defective Product + Damaged in Transit** account for 82% of total return value ($15,750 out of $19,270)
- **Customers are largely not at fault** — "Customer Ordered Wrong Part" is only $450 of total returns, pointing to a supply chain and quality issue

---

## 🛠️ Skills Demonstrated

- IF & Nested IF formulas
- VLOOKUP across multiple tables
- TEXT, MONTH, ROUNDUP functions
- Pivot Tables & Pivot Charts
- Calculated Fields
- Timeline Slicer connected to multiple pivots
- Slicers connected to multiple pivot tables
- KPI dashboard design
- Dark theme professional formatting

---

## 🚀 How to Use

1. Download `Sales_MIS_Dashboard_Excel.xlsx`
2. Open in Microsoft Excel
3. Go to the **DashBoard** sheet
4. Use the **slicers** and **timeline** to filter data interactively
5. Explore individual pivot sheets for detailed analysis

---

## 👤 Author
**Ajmal**
- 📧 *(patelajmal04@gmail.com)*
- 💼 *(www.linkedin.com/in/mohammed-ajmal-patel)*
- 🐙 *(https://ajmalpatel16.github.io/Portfolio/)*
