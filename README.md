# SuperStore Sales Dashboard | Power BI

Interactive Power BI dashboard built from an Excel order-line dataset to explore sales, profit, shipping, geography and product mix.

# [Dashboard preview]<img width="1211" height="676" alt="dashboard" src="https://github.com/user-attachments/assets/8b9345c1-3413-45de-a858-56b27273c5c3" />



## Business problem
Sales teams need a single interactive report to monitor revenue, profit, order volume, delivery time and performance across regions and product segments.

## Dataset and analysis
- Excel source: **5,901 order-line records**, **23 fields**.
- Distinct Order IDs: **3,003**; row count is not order count.
- Raw source sales: **1,565,804.32** (currency unspecified).
- Raw source profit: **175,262.11**; profit margin: **11.19%**.
- Raw quantity: **22,317** units.
- Dates: 2019, 2020.
- **0 rows** have Ship Date earlier than Order Date; investigate before calculating delivery duration.

> The screenshot shows rounded card labels of 1.6M sales, 22K orders, 175K profit and 4 average delivery days. Those report values should not be treated as independently validated source totals without checking PBIX model filters, transformations and measure definitions.

## Dashboard visuals
- KPI cards: sales, orders, profit, average delivery days
- Monthly sales and profit by year
- Sales by payment mode, region and customer segment
- Sales by ship mode, category and subcategory
- State-level map with sales/profit context
- Region slicer for cross-filtering

## Data dictionary
| Field | Description |
|---|---|
| Row ID+O6G3A1:R6 | Row-level identifier; header appears corrupted or concatenated in the supplied file. |
| Order ID | Order identifier; multiple product rows can share one order. |
| Order Date | Date of order placement; Excel serial date. |
| Ship Date | Recorded shipping date; verify against order date. |
| Ship Mode | Shipping service class. |
| Customer ID | Customer identifier; pseudonymize before publication. |
| Customer Name | Customer name; remove from public sample. |
| Segment | Consumer, Corporate or Home Office customer group. |
| Country | Country of sale. |
| City | Customer city. |
| State | US state used for mapping. |
| Region | Central, East, South or West. |
| Product ID | Product identifier. |
| Category | High-level product family. |
| Sub-Category | Detailed product group. |
| Product Name | Descriptive product label. |
| Sales | Sales amount on an order line; currency not documented. |
| Quantity | Units on an order line. |
| Profit | Reported profit on an order line; source calculation not specified. |
| Returns | Return flag/status; contains #N/A and requires normalization. |
| Payment Mode | Payment channel, including COD, Online and Cards. |
| ind1 / ind2 | Additional fields; source semantics not documented. |

## DAX measures
```dax
Total Sales = SUM ( SuperStore[Sales] )
Total Profit = SUM ( SuperStore[Profit] )
Total Orders = DISTINCTCOUNT ( SuperStore[Order ID] )
Total Quantity = SUM ( SuperStore[Quantity] )
Profit Margin % = DIVIDE ( [Total Profit], [Total Sales] )
```

## Rebuild instructions
See `SuperStore_Rebuild_Guide.pdf` for import, cleaning, formulas, visuals, slicer interactions, validation, Power BI refresh and optional Excel VBA examples. Power BI does not use VBA.

## Files
- `SuperStore Sales Dashboard.pbix` — Power BI report
- `SuperStore_Sales_Dataset.xlsx` — Excel source (review privacy/license before publishing)
- `dashboard.png` — report preview
- `SuperStore_Project_Documentation.pdf` — business problem and terminology
- `SuperStore_Rebuild_Guide.pdf` — detailed tutorial

## Limitations
- Dataset contains customer names; anonymize before sharing publicly.
- Order lines repeat Order IDs, so distinct order count is required.
- Shipping date inconsistencies can distort delivery averages.
- Source Profit definition and currency are not documented.
- The first header is malformed; `ind1`/`ind2` are undefined.
- VBA snippets in the tutorial are optional Excel examples, not Power BI macros.
