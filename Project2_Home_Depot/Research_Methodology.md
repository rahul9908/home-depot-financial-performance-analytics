# Research methodology

## Research question

How did Home Depot's revenue composition, operating economics, working capital, capital efficiency, and cash-conversion quality change from fiscal 2021 through fiscal 2025, and what operating assumptions would be required for improvement through fiscal 2028?

## Analytical design

The project uses a longitudinal single-company case study. Audited financial statements are standardized into consistent line items and evaluated through common-size analysis, growth rates, liquidity and leverage ratios, cash conversion, working-capital efficiency, return on invested capital, and operating-income attribution.

The analysis distinguishes three types of evidence:

1. **Reported outcomes:** audited consolidated financial-statement amounts.
2. **Accounting decompositions:** identities that reconcile changes in operating income or working capital.
3. **Conditional forecasts:** results that follow from stated assumptions and are not statistical predictions.

## Key definitions

- NOPAT = Operating Income × (1 − Effective Tax Rate).
- Invested Capital = Current Debt + Long-term Debt + Stockholders' Equity − Cash.
- ROIC = NOPAT ÷ Average Invested Capital.
- Accrual Ratio = (Net Income − Operating Cash Flow) ÷ Average Total Assets.
- Net Operating Working Capital = Accounts Receivable + Inventory − Accounts Payable.
- Cash Conversion Cycle = Receivable Days + Inventory Days − Payable Days.
- Free Cash Flow = Operating Cash Flow − Capital Expenditures.

## Operating-income attribution

Year-over-year operating-income change is reconciled as:

1. Revenue effect = Change in Revenue × Prior-year Gross Margin.
2. Gross-margin rate effect = Current Revenue × Change in Gross Margin.
3. SG&A effect = −Change in SG&A.
4. D&A effect = −Change in D&A Expense.

The four effects sum to the reported change in operating income. This is an accounting bridge, not a causal model.

## Scenario methodology

Downside, base, and upside cases vary revenue growth, gross margin, operating margin, CFO conversion, and capex intensity. Interest expense is held flat and the effective tax rate is stabilized. A separate two-driver sensitivity tests fiscal 2028 operating income across revenue CAGR and operating-margin assumptions.

## Validity and limitations

- Internal validity is limited by acquisitions, a 53-week fiscal year, inflation, mix changes, and aggregation across geographies and customer groups.
- External validity is limited because this is a single-company study in home-improvement retail.
- Five annual observations support trend and decomposition analysis but not reliable causal regression.
- SEC XBRL comparability depends on consistent tagging. The transformation script selects annual facts by exact fiscal-period end date and reconciles key statement identities.
- Forecast conclusions should be read as conditional statements, not probability-weighted estimates or investment advice.

## Reproducibility

The Python transformation retrieves the SEC Company Facts feed, standardizes line items, calculates ratios, and exports long-form CSV files. The Excel workbook links outputs to source data and includes terminal audit checks. Power BI queries consume the same CSV outputs.
