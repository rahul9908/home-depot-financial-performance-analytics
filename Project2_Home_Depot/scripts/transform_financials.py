"""Build a five-year Home Depot financial-analysis dataset from SEC Company Facts.

Outputs are expressed in USD millions and are designed for Excel and Power BI.
The script deliberately selects facts by fiscal period end date rather than SEC
``fy`` because comparative facts are repeated in later filings.
"""

from __future__ import annotations

import csv
import json
import math
import urllib.request
from datetime import date
from pathlib import Path


CIK = "0000354950"
COMPANY = "The Home Depot, Inc."
SEC_FACTS_URL = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"
SOURCE_10K = "https://www.sec.gov/Archives/edgar/data/354950/000162828026019436/hd-20260201.htm"
USER_AGENT = "Public-company financial analysis portfolio contact@example.com"
ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"

PERIODS = {
    2021: "2022-01-30",
    2022: "2023-01-29",
    2023: "2024-01-28",
    2024: "2025-02-02",
    2025: "2026-02-01",
}


def fetch_json(url: str) -> dict:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.load(response)


def facts_for(data: dict, tag: str) -> list[dict]:
    fact = data["facts"]["us-gaap"].get(tag)
    if not fact:
        return []
    units = fact.get("units", {})
    return units.get("USD", next(iter(units.values()), []))


def annual_value(data: dict, tag: str, end: str, instant: bool = False) -> float | None:
    candidates = []
    for item in facts_for(data, tag):
        if item.get("end") != end or item.get("form") != "10-K":
            continue
        if instant:
            score = (item.get("filed", ""), item.get("accn", ""))
        else:
            start = item.get("start")
            if not start:
                continue
            duration = (date.fromisoformat(end) - date.fromisoformat(start)).days
            if not 330 <= duration <= 380:
                continue
            score = (duration, item.get("filed", ""), item.get("accn", ""))
        candidates.append((score, item["val"]))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1] / 1_000_000


def first_available(data: dict, tags: list[str], end: str, instant: bool = False) -> float | None:
    for tag in tags:
        value = annual_value(data, tag, end, instant)
        if value is not None:
            return value
    return None


def safe_div(a: float | None, b: float | None) -> float | None:
    return None if a is None or b in (None, 0) else a / b


def write_csv(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    data = fetch_json(SEC_FACTS_URL)
    raw_rows: list[dict] = []
    wide: dict[int, dict[str, float | None]] = {}

    definitions = [
        ("Income Statement", "Revenue", ["RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet"], False),
        ("Income Statement", "Cost of Sales", ["CostOfRevenue"], False),
        ("Income Statement", "Gross Profit", ["GrossProfit"], False),
        ("Income Statement", "SG&A", ["SellingGeneralAndAdministrativeExpense"], False),
        ("Income Statement", "Depreciation & Amortization Expense", ["DepreciationAndAmortization"], False),
        ("Income Statement", "Operating Income", ["OperatingIncomeLoss"], False),
        ("Income Statement", "Interest Expense", ["InterestExpenseNonoperating", "InterestExpense"], False),
        ("Income Statement", "Income Before Tax", ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"], False),
        ("Income Statement", "Income Tax Expense", ["IncomeTaxExpenseBenefit"], False),
        ("Income Statement", "Net Income", ["NetIncomeLoss"], False),
        ("Balance Sheet", "Cash & Equivalents", ["CashAndCashEquivalentsAtCarryingValue"], True),
        ("Balance Sheet", "Accounts Receivable", ["AccountsReceivableNetCurrent"], True),
        ("Balance Sheet", "Inventory", ["InventoryNet"], True),
        ("Balance Sheet", "Current Assets", ["AssetsCurrent"], True),
        ("Balance Sheet", "Total Assets", ["Assets"], True),
        ("Balance Sheet", "Accounts Payable", ["AccountsPayableCurrent"], True),
        ("Balance Sheet", "Current Liabilities", ["LiabilitiesCurrent"], True),
        ("Balance Sheet", "Current Debt", ["ShortTermBorrowings", "LongTermDebtAndCapitalLeaseObligationsCurrent", "LongTermDebtCurrent"], True),
        ("Balance Sheet", "Long-term Debt", ["LongTermDebtAndCapitalLeaseObligations", "LongTermDebt", "OtherLongTermDebtNoncurrent"], True),
        ("Balance Sheet", "Operating Lease Liabilities", ["OperatingLeaseLiability"], True),
        ("Balance Sheet", "Total Liabilities", ["Liabilities"], True),
        ("Balance Sheet", "Stockholders' Equity", ["StockholdersEquity"], True),
        ("Cash Flow", "Operating Cash Flow", ["NetCashProvidedByUsedInOperatingActivities"], False),
        ("Cash Flow", "D&A Add-back", ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"], False),
        ("Cash Flow", "Change in Receivables", ["IncreaseDecreaseInReceivables"], False),
        ("Cash Flow", "Change in Inventory", ["IncreaseDecreaseInInventories"], False),
        ("Cash Flow", "Change in A/P & Accrued Liabilities", ["IncreaseDecreaseInAccountsPayableAndAccruedLiabilities"], False),
        ("Cash Flow", "Capital Expenditures", ["PaymentsToAcquireProductiveAssets"], False),
        ("Cash Flow", "Business Acquisitions", ["PaymentsToAcquireBusinessesNetOfCashAcquired"], False),
        ("Cash Flow", "Investing Cash Flow", ["NetCashProvidedByUsedInInvestingActivities"], False),
        ("Cash Flow", "Share Repurchases", ["PaymentsForRepurchaseOfCommonStock"], False),
        ("Cash Flow", "Dividends Paid", ["PaymentsOfDividendsCommonStock", "DividendsCommonStockCash"], False),
        ("Cash Flow", "Financing Cash Flow", ["NetCashProvidedByUsedInFinancingActivities"], False),
    ]

    for fiscal_year, end in PERIODS.items():
        row: dict[str, float | None] = {}
        for statement, line_item, tags, instant in definitions:
            value = first_available(data, tags, end, instant)
            row[line_item] = value
            raw_rows.append({
                "Company": COMPANY,
                "Ticker": "HD",
                "Fiscal Year": fiscal_year,
                "Fiscal Year End": end,
                "Statement": statement,
                "Line Item": line_item,
                "Value USD millions": "" if value is None else round(value, 3),
                "Source": SOURCE_10K if fiscal_year >= 2023 else "SEC Company Facts comparative 10-K data",
            })
        wide[fiscal_year] = row

    ratio_rows: list[dict] = []
    for year, row in wide.items():
        previous = wide.get(year - 1, {})
        revenue = row.get("Revenue")
        cogs = row.get("Cost of Sales")
        current_debt = row.get("Current Debt") or 0
        long_debt = row.get("Long-term Debt") or 0
        ratios = {
            "Revenue Growth": safe_div(revenue, previous.get("Revenue")) - 1 if previous.get("Revenue") else None,
            "Gross Margin": safe_div(row.get("Gross Profit"), revenue),
            "Operating Margin": safe_div(row.get("Operating Income"), revenue),
            "Net Margin": safe_div(row.get("Net Income"), revenue),
            "SG&A as % Revenue": safe_div(row.get("SG&A"), revenue),
            "Current Ratio": safe_div(row.get("Current Assets"), row.get("Current Liabilities")),
            "Quick Ratio": safe_div((row.get("Cash & Equivalents") or 0) + (row.get("Accounts Receivable") or 0), row.get("Current Liabilities")),
            "Debt / Equity": safe_div(current_debt + long_debt, row.get("Stockholders' Equity")),
            "Debt / Assets": safe_div(current_debt + long_debt, row.get("Total Assets")),
            "CFO / Net Income": safe_div(row.get("Operating Cash Flow"), row.get("Net Income")),
            "Free Cash Flow Margin": safe_div((row.get("Operating Cash Flow") or 0) - (row.get("Capital Expenditures") or 0), revenue),
            "Inventory Turnover": safe_div(cogs, ((row.get("Inventory") or 0) + (previous.get("Inventory") or row.get("Inventory") or 0)) / 2),
            "Receivable Days": safe_div(((row.get("Accounts Receivable") or 0) + (previous.get("Accounts Receivable") or row.get("Accounts Receivable") or 0)) / 2 * 365, revenue),
            "Inventory Days": safe_div(((row.get("Inventory") or 0) + (previous.get("Inventory") or row.get("Inventory") or 0)) / 2 * 365, cogs),
            "Payable Days": safe_div(((row.get("Accounts Payable") or 0) + (previous.get("Accounts Payable") or row.get("Accounts Payable") or 0)) / 2 * 365, cogs),
        }
        for metric, value in ratios.items():
            ratio_rows.append({"Company": COMPANY, "Ticker": "HD", "Fiscal Year": year, "Metric": metric, "Value": "" if value is None else round(value, 6)})

    # Scenario assumptions use fiscal 2025 as the common starting point.
    scenarios = {
        "Downside": {"Revenue Growth": [0.00, 0.01, 0.02], "Gross Margin": [0.330, 0.331, 0.332], "Operating Margin": [0.120, 0.122, 0.124], "CFO / Net Income": [1.05, 1.08, 1.10]},
        "Base": {"Revenue Growth": [0.025, 0.030, 0.035], "Gross Margin": [0.334, 0.335, 0.336], "Operating Margin": [0.128, 0.130, 0.132], "CFO / Net Income": [1.12, 1.15, 1.17]},
        "Upside": {"Revenue Growth": [0.050, 0.055, 0.050], "Gross Margin": [0.338, 0.340, 0.341], "Operating Margin": [0.135, 0.139, 0.142], "CFO / Net Income": [1.18, 1.20, 1.22]},
    }
    scenario_rows: list[dict] = []
    base_revenue = wide[2025]["Revenue"] or 0
    tax_rate = safe_div(wide[2025]["Income Tax Expense"], wide[2025]["Income Before Tax"]) or 0.235
    interest = wide[2025]["Interest Expense"] or 2_500
    for scenario, assumptions in scenarios.items():
        revenue = base_revenue
        for index, forecast_year in enumerate((2026, 2027, 2028)):
            revenue *= 1 + assumptions["Revenue Growth"][index]
            gross_profit = revenue * assumptions["Gross Margin"][index]
            operating_income = revenue * assumptions["Operating Margin"][index]
            pretax = operating_income - interest
            net_income = pretax * (1 - tax_rate)
            cfo = net_income * assumptions["CFO / Net Income"][index]
            for metric, value in {
                "Revenue": revenue,
                "Gross Profit": gross_profit,
                "Operating Income": operating_income,
                "Net Income": net_income,
                "Operating Cash Flow": cfo,
            }.items():
                scenario_rows.append({"Scenario": scenario, "Fiscal Year": forecast_year, "Metric": metric, "Value USD millions": round(value, 1)})
            for metric in ("Revenue Growth", "Gross Margin", "Operating Margin", "CFO / Net Income"):
                scenario_rows.append({"Scenario": scenario, "Fiscal Year": forecast_year, "Metric": metric, "Value USD millions": assumptions[metric][index]})

    write_csv(DATA_DIR / "financials_long.csv", raw_rows, list(raw_rows[0]))
    write_csv(DATA_DIR / "ratios_long.csv", ratio_rows, list(ratio_rows[0]))
    write_csv(DATA_DIR / "scenario_outlook.csv", scenario_rows, list(scenario_rows[0]))
    (DATA_DIR / "model_tables.json").write_text(
        json.dumps({"financials": raw_rows, "ratios": ratio_rows, "scenarios": scenario_rows}, indent=2),
        encoding="utf-8",
    )
    (DATA_DIR / "sec_companyfacts_snapshot.json").write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"Wrote {len(raw_rows)} financial rows, {len(ratio_rows)} ratio rows, and {len(scenario_rows)} scenario rows to {DATA_DIR}")


if __name__ == "__main__":
    main()
