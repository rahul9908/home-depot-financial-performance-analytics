# Power BI dashboard build guide

## Model

Load `financials_long.csv`, `ratios_long.csv`, and `scenario_outlook.csv` with the supplied Power Query script. Create a one-column fiscal-year dimension containing 2021–2028 and relate it one-to-many to each fact table. Keep the source tables in long format.

## Page 1: Executive performance

- Cards: Revenue, Operating Margin, Operating Cash Flow, CFO Conversion, Current Ratio, Free Cash Flow.
- Combo chart: Revenue columns and Operating Margin line by fiscal year.
- Line chart: Net Income and Operating Cash Flow by fiscal year.
- Clustered bar: Gross Margin, Operating Margin, and Net Margin by fiscal year.
- Matrix: selected income statement and cash-flow line items by fiscal year.
- Slicers: fiscal year and statement.

## Page 2: Liquidity and working capital

- Lines: Current Ratio and Quick Ratio.
- Lines: Receivable Days, Inventory Days, Payable Days, and Cash Conversion Cycle.
- Columns: Accounts Receivable, Inventory, and Accounts Payable.
- Cards: Debt / Assets, Debt / Equity, and FY2025 net operating working capital.

## Page 3: Scenario outlook

- Slicer: Scenario.
- Lines: Revenue and Operating Cash Flow for fiscal 2026–2028.
- Cards: fiscal 2028 Revenue, Operating Income, Net Income, and Operating Cash Flow.
- Matrix: Metrics by fiscal year and scenario.

## Page 4: Capital efficiency and earnings quality

- Lines: ROIC, ROA, and asset turnover.
- Lines: CFO conversion, FCF-to-net-income conversion, and dividend coverage.
- Columns: NOPAT and invested capital.
- Table: accrual ratio, capex-to-D&A, interest coverage, and NOWC-to-revenue.

## Page 5: Advanced drivers

- Waterfall-style columns or matrix: revenue effect, gross-margin effect, SG&A effect, and D&A effect on operating-income change.
- Columns: Primary segment revenue and Other segment revenue for fiscal 2023–2025.
- Matrix: Building Materials, Décor, and Hardlines revenue.
- Decision-threshold panel: positive Primary segment growth, operating margin above 13%, CFO conversion above 1.0x, and improving ROIC.

## Formatting

Use the supplied theme. Use dark blue for historical results, light blue for forecasts, green for cash-flow measures, and red only for adverse variance. Display financials in USD billions or millions consistently and percentages with one decimal.

## Refresh

Run `scripts/transform_financials.py` first, then refresh Power BI. The script downloads a fresh SEC Company Facts snapshot and rewrites all three CSVs. Review SEC tag changes before relying on a new period.
