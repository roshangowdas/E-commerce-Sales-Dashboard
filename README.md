# README – E-commerce Sales Dashboard (Power BI)

## Project Title
**E-commerce Sales Dashboard using Power BI Desktop**

---

## Project Overview
This project demonstrates the creation of an interactive E-commerce Sales Dashboard in Microsoft Power BI Desktop using data from:

- An E-commerce Sales Excel dataset
- A State Logistics CSV dataset

The dashboard provides insights into:

- Sales performance
- Profitability
- Customer distribution
- Product performance
- Shipment analysis
- Regional sales trends
- Year-over-Year (YoY) growth

The dashboard is designed with a dark professional theme, interactive slicers, KPI cards, geographic mapping, and advanced DAX calculations.

---

## Files Included

| File Name | Description |
|---|---|
| `EcommerceSalesDashboard.pbix` | Main Power BI Dashboard file |
| `Ecommerce_Sales.xlsx` | E-commerce transactional dataset |
| `State_Logistics.csv` | State logistics/location dataset |
| `README.md` | Project documentation |

---

## Tools & Technologies Used

- Microsoft Power BI Desktop
- DAX (Data Analysis Expressions)
- Excel Data Source
- CSV Data Source
- Data Modeling
- Interactive Visualizations

---

# Dashboard Features

## 1. Data Integration & Modeling

### Imported Data Sources
- E-commerce Excel dataset
- State logistics CSV dataset

### Relationships Created
- `State Logistics[State]` → `E-commerce[State]`
- Relationship Type: One-to-Many
- Cross-filter Direction: Single

### Calendar Table
A custom Calendar table was created using DAX for time intelligence analysis.

```DAX
Calendar =
VAR MinDate = MIN('E-commerce'[Order Date])
VAR MaxDate = MAX('E-commerce'[Order Date])

RETURN
ADDCOLUMNS(
    CALENDAR(MinDate, MaxDate),
    "Year", YEAR([Date]),
    "Month", FORMAT([Date], "MMMM"),
    "Month Number", MONTH([Date]),
    "Day", DAY([Date]),
    "Year-Month", FORMAT([Date], "YYYY-MM"),
    "Quarter", "Q" & FORMAT([Date], "Q")
)
