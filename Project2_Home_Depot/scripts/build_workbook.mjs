import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectDir = path.resolve("Project2_Home_Depot");
const outputDir = path.resolve("outputs", "project2_home_depot");
const model = JSON.parse(await fs.readFile(path.join(projectDir, "data", "model_tables.json"), "utf8"));

const wb = Workbook.create();
const dashboard = wb.worksheets.add("Dashboard");
const assumptions = wb.worksheets.add("Assumptions");
const statements = wb.worksheets.add("Statements");
const ratios = wb.worksheets.add("Ratios & WC");
const advanced = wb.worksheets.add("Advanced Analysis");
const scenario = wb.worksheets.add("Scenario Outlook");
const sensitivity = wb.worksheets.add("Sensitivity");
const data = wb.worksheets.add("Data");
const audit = wb.worksheets.add("Audit");

const navy = "#17365D";
const blue = "#2F75B5";
const lightBlue = "#D9EAF7";
const paleBlue = "#EAF2F8";
const amber = "#FFF2CC";
const green = "#548235";
const lightGreen = "#E2F0D9";
const red = "#C00000";
const gray = "#E7E6E6";
const dark = "#1F1F1F";
const font = "Arial";
const moneyFmt = '$#,##0;[Red]($#,##0);-';
const pctFmt = '0.0%;[Red](0.0%);-';
const ratioFmt = '0.00x;[Red](0.00x);-';

for (const sheet of [dashboard, assumptions, statements, ratios, advanced, scenario, sensitivity, data, audit]) {
  sheet.showGridLines = false;
  sheet.getRange("A1:Z100").format.font = { name: font, size: 10, color: dark };
}
dashboard.tabColor = navy;
assumptions.tabColor = blue;
data.tabColor = "#D9C7A5";
audit.tabColor = "#A6A6A6";

function title(sheet, text, note = "") {
  sheet.getRange("C2:N2").merge();
  sheet.getRange("C2").values = [[text]];
  sheet.getRange("C2").format.font = { name: font, size: 16, bold: true, color: navy };
  sheet.getRange("C3:N3").format.borders = { bottom: { style: "thin", color: blue } };
  if (note) {
    sheet.getRange("C4:N4").merge();
    sheet.getRange("C4").values = [[note]];
    sheet.getRange("C4").format.font = { name: font, size: 10, italic: true, color: "#666666" };
  }
}

function section(sheet, rangeAddress, text) {
  const r = sheet.getRange(rangeAddress);
  r.format.fill = lightBlue;
  r.format.font = { name: font, size: 10, bold: true, color: navy };
  r.format.borders = { preset: "outside", style: "thin", color: "#A9C4DD" };
  r.getCell(0, 0).values = [[text]];
}

function header(r) {
  r.format.fill = navy;
  r.format.font = { name: font, size: 10, bold: true, color: "#FFFFFF" };
  r.format.horizontalAlignment = "center";
  r.format.verticalAlignment = "center";
  r.format.borders = { insideVertical: { style: "thin", color: "#FFFFFF" }, bottom: { style: "thin", color: navy } };
}

function setWidths(sheet) {
  sheet.getRange("A:A").format.columnWidth = 3;
  sheet.getRange("B:B").format.columnWidth = 3;
  sheet.getRange("C:C").format.columnWidth = 29;
}

// Data sheet: standardized source table and wide input block.
title(data, "Source financial data", "USD millions. Fiscal years follow Home Depot's 52/53-week calendar.");
data.getRange("C6:J6").values = [["Company", "Ticker", "Fiscal Year", "Fiscal Year End", "Statement", "Line Item", "Value USD millions", "Source"]];
header(data.getRange("C6:J6"));
const finRows = model.financials.map(r => [r.Company, r.Ticker, r["Fiscal Year"], new Date(`${r["Fiscal Year End"]}T00:00:00`), r.Statement, r["Line Item"], r["Value USD millions"] === "" ? null : r["Value USD millions"], r.Source]);
data.getRange("C7").write(finRows);
data.getRange(`F7:F${6 + finRows.length}`).format.numberFormat = "yyyy-mm-dd";
data.getRange(`I7:I${6 + finRows.length}`).format.numberFormat = moneyFmt;
data.freezePanes.freezeRows(6);
data.getRange("C:C").format.columnWidth = 23;
data.getRange("D:D").format.columnWidth = 9;
data.getRange("E:F").format.columnWidth = 13;
data.getRange("G:G").format.columnWidth = 18;
data.getRange("H:H").format.columnWidth = 34;
data.getRange("I:I").format.columnWidth = 18;
data.getRange("J:J").format.columnWidth = 42;
data.getRange("L5:R5").values = [["Line Item", 2021, 2022, 2023, 2024, 2025, "Source URL"]];
header(data.getRange("L5:R5"));
const lineItems = [...new Set(model.financials.map(r => r["Line Item"]))];
const wideRows = lineItems.map(item => {
  const values = [item];
  for (const y of [2021, 2022, 2023, 2024, 2025]) {
    const match = model.financials.find(r => r["Line Item"] === item && r["Fiscal Year"] === y);
    values.push(match && match["Value USD millions"] !== "" ? match["Value USD millions"] : null);
  }
  values.push("https://www.sec.gov/Archives/edgar/data/354950/000162828026019436/hd-20260201.htm");
  return values;
});
data.getRange("L6").write(wideRows);
data.getRange(`M6:Q${5 + wideRows.length}`).format.numberFormat = moneyFmt;
data.getRange("L:L").format.columnWidth = 38;
data.getRange("M:Q").format.columnWidth = 12;
data.getRange("R:R").format.columnWidth = 52;

const dataRow = Object.fromEntries(lineItems.map((x, i) => [x, i + 6]));
const years = [2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028];
const cols = ["D", "E", "F", "G", "H", "I", "J", "K"];

// Assumptions with one authoritative selector.
title(assumptions, "Forecast assumptions", "Select one case. Blue cells are editable inputs; green text is the active selection.");
assumptions.getRange("C6:D6").values = [["Case selector", 2]];
assumptions.getRange("C7:D9").values = [["1", "Downside"], ["2", "Base"], ["3", "Upside"]];
assumptions.getRange("D6").dataValidation = { rule: { type: "list", values: [1, 2, 3] } };
assumptions.getRange("D6").format.fill = amber;
assumptions.getRange("D6").format.font = { name: font, size: 11, bold: true, color: "#0000FF" };
assumptions.getRange("F6:H6").values = [[2026, 2027, 2028]];
header(assumptions.getRange("F6:H6"));
const driverBlocks = [
  { label: "Revenue Growth", start: 11, values: [[0.00,0.01,0.02],[0.025,0.03,0.035],[0.05,0.055,0.05]], fmt: pctFmt },
  { label: "Gross Margin", start: 17, values: [[0.33,0.331,0.332],[0.334,0.335,0.336],[0.338,0.34,0.341]], fmt: pctFmt },
  { label: "Operating Margin", start: 23, values: [[0.12,0.122,0.124],[0.128,0.13,0.132],[0.135,0.139,0.142]], fmt: pctFmt },
  { label: "CFO / Net Income", start: 29, values: [[1.05,1.08,1.10],[1.12,1.15,1.17],[1.18,1.20,1.22]], fmt: ratioFmt },
  { label: "Capex / Revenue", start: 35, values: [[0.025,0.025,0.025],[0.023,0.023,0.023],[0.022,0.022,0.022]], fmt: pctFmt },
];
for (const block of driverBlocks) {
  section(assumptions, `C${block.start}:H${block.start}`, block.label);
  assumptions.getRange(`C${block.start+1}:E${block.start+4}`).values = [
    ["Active Selection", null, null],
    ["Downside", null, null],
    ["Base", null, null],
    ["Upside", null, null],
  ];
  assumptions.getRange(`F${block.start+2}:H${block.start+4}`).values = block.values;
  assumptions.getRange(`F${block.start+1}`).formulas = [[`=CHOOSE($D$6,F${block.start+2},F${block.start+3},F${block.start+4})`]];
  assumptions.getRange(`F${block.start+1}:H${block.start+1}`).fillRight();
  assumptions.getRange(`F${block.start+1}:H${block.start+4}`).format.numberFormat = block.fmt;
  assumptions.getRange(`F${block.start+2}:H${block.start+4}`).format.fill = amber;
  assumptions.getRange(`F${block.start+2}:H${block.start+4}`).format.font = { name: font, size: 10, color: "#0000FF" };
  assumptions.getRange(`C${block.start+1}:H${block.start+1}`).format.font = { name: font, size: 10, bold: true, color: green };
}
assumptions.getRange("C41:F44").values = [["Other forecast inputs", "Value", "Basis", "Notes"], ["Interest Expense", 2412, "FY2025 actual", "Held flat"], ["Tax Rate", 0.239, "FY2025 effective rate", "Applied to pretax income"], ["D&A / Revenue", 0.02, "Rounded FY2025", "Used to estimate operating expenses"]];
header(assumptions.getRange("C41:F41"));
assumptions.getRange("D42:D44").format.fill = amber;
assumptions.getRange("D42:D44").format.font = { name: font, size: 10, color: "#0000FF" };
assumptions.getRange("D42").format.numberFormat = moneyFmt;
assumptions.getRange("D43:D44").format.numberFormat = pctFmt;
setWidths(assumptions);
assumptions.getRange("D:D").format.columnWidth = 15;
assumptions.getRange("E:E").format.columnWidth = 3;
assumptions.getRange("F:H").format.columnWidth = 13;

// Three-statement analysis and active forecast.
title(statements, "Financial statement analysis", "Actuals through fiscal 2025; forecast updates with the selected case.");
statements.getRange("C5:K5").values = [["USD millions", ...years]];
header(statements.getRange("C5:K5"));
statements.getRange("C6:D6").values = [["Case Selected:", null]];
statements.getRange("D6").formulas = [["=CHOOSE('Assumptions'!$D$6,\"Downside\",\"Base\",\"Upside\")"]];
statements.getRange("D6").format.font = { name: font, size: 10, bold: true, color: green };
const statementRows = {
  "Revenue": 9, "Revenue Growth": 10, "Cost of Sales": 11, "Gross Profit": 12, "Gross Margin": 13,
  "SG&A": 15, "SG&A as % Revenue": 16, "Depreciation & Amortization Expense": 17, "Operating Income": 18,
  "Operating Margin": 19, "Interest Expense": 21, "Income Before Tax": 22, "Income Tax Expense": 23,
  "Net Income": 24, "Net Margin": 25, "Operating Cash Flow": 29, "CFO / Net Income": 30,
  "Capital Expenditures": 31, "Free Cash Flow": 32,
  "Cash & Equivalents": 36, "Accounts Receivable": 37, "Inventory": 38, "Current Assets": 39,
  "Total Assets": 40, "Accounts Payable": 42, "Current Debt": 43, "Current Liabilities": 44,
  "Long-term Debt": 45, "Operating Lease Liabilities": 46, "Total Liabilities": 47, "Stockholders' Equity": 48,
};
section(statements, "C8:K8", "Income statement");
section(statements, "C28:K28", "Cash flow");
section(statements, "C35:K35", "Balance sheet");
for (const [label, row] of Object.entries(statementRows)) statements.getRange(`C${row}`).values = [[label]];

const historicalItems = Object.keys(statementRows).filter(x => !x.includes("Margin") && !x.includes("Growth") && x !== "CFO / Net Income" && x !== "Free Cash Flow");
for (const item of historicalItems) {
  const row = statementRows[item];
  const sourceRow = dataRow[item];
  if (!sourceRow) continue;
  statements.getRange(`D${row}`).formulas = [[`='Data'!M${sourceRow}`]];
  statements.getRange(`D${row}:H${row}`).fillRight();
}
for (let c = 0; c < 5; c++) {
  const col = cols[c];
  const prev = c === 0 ? null : cols[c-1];
  statements.getRange(`${col}10`).formulas = [[prev ? `=${col}9/${prev}9-1` : '=""']];
  statements.getRange(`${col}13`).formulas = [[`=${col}12/${col}9`]];
  statements.getRange(`${col}16`).formulas = [[`=${col}15/${col}9`]];
  statements.getRange(`${col}19`).formulas = [[`=${col}18/${col}9`]];
  statements.getRange(`${col}25`).formulas = [[`=${col}24/${col}9`]];
  statements.getRange(`${col}30`).formulas = [[`=${col}29/${col}24`]];
  statements.getRange(`${col}32`).formulas = [[`=${col}29-${col}31`]];
}
for (let c = 5; c < 8; c++) {
  const col = cols[c];
  const prev = cols[c-1];
  const ac = ["F", "G", "H"][c-5];
  statements.getRange(`${col}10`).formulas = [[`='Assumptions'!${ac}12`]];
  statements.getRange(`${col}9`).formulas = [[`=${prev}9*(1+${col}10)`]];
  statements.getRange(`${col}13`).formulas = [[`='Assumptions'!${ac}18`]];
  statements.getRange(`${col}12`).formulas = [[`=${col}9*${col}13`]];
  statements.getRange(`${col}11`).formulas = [[`=${col}9-${col}12`]];
  statements.getRange(`${col}17`).formulas = [[`=${col}9*'Assumptions'!$D$44`]];
  statements.getRange(`${col}19`).formulas = [[`='Assumptions'!${ac}24`]];
  statements.getRange(`${col}18`).formulas = [[`=${col}9*${col}19`]];
  statements.getRange(`${col}15`).formulas = [[`=${col}12-${col}17-${col}18`]];
  statements.getRange(`${col}16`).formulas = [[`=${col}15/${col}9`]];
  statements.getRange(`${col}21`).formulas = [["='Assumptions'!$D$42"]];
  statements.getRange(`${col}22`).formulas = [[`=${col}18-${col}21`]];
  statements.getRange(`${col}23`).formulas = [[`=${col}22*'Assumptions'!$D$43`]];
  statements.getRange(`${col}24`).formulas = [[`=${col}22-${col}23`]];
  statements.getRange(`${col}25`).formulas = [[`=${col}24/${col}9`]];
  statements.getRange(`${col}30`).formulas = [[`='Assumptions'!${ac}30`]];
  statements.getRange(`${col}29`).formulas = [[`=${col}24*${col}30`]];
  statements.getRange(`${col}31`).formulas = [[`=${col}9*'Assumptions'!${ac}36`]];
  statements.getRange(`${col}32`).formulas = [[`=${col}29-${col}31`]];
}
statements.getRange("D9:K48").format.numberFormat = moneyFmt;
for (const row of [10,13,16,19,25]) statements.getRange(`D${row}:K${row}`).format.numberFormat = pctFmt;
statements.getRange("D30:K30").format.numberFormat = ratioFmt;
for (const row of [10,13,16,19,25,30]) statements.getRange(`C${row}:K${row}`).format.font = { name: font, size: 10, italic: true, color: dark };
for (const row of [12,18,24,32,39,40,44,47,48]) statements.getRange(`C${row}:K${row}`).format.font = { name: font, size: 10, bold: true, color: dark };
statements.getRange("I5:K48").format.fill = paleBlue;
statements.freezePanes.freezeRows(5);
statements.freezePanes.freezeColumns(3);
setWidths(statements);
statements.getRange("D:K").format.columnWidth = 13;

// Ratios and working-capital analysis.
title(ratios, "Ratios and working capital", "Liquidity, leverage, cash conversion, and efficiency trends.");
ratios.getRange("C6:H6").values = [["Metric", 2021, 2022, 2023, 2024, 2025]];
header(ratios.getRange("C6:H6"));
const ratioNames = ["Revenue Growth","Gross Margin","Operating Margin","Net Margin","SG&A as % Revenue","Current Ratio","Quick Ratio","Debt / Equity","Debt / Assets","CFO / Net Income","Free Cash Flow Margin","Inventory Turnover","Receivable Days","Inventory Days","Payable Days","Cash Conversion Cycle"];
ratioNames.forEach((x,i)=>ratios.getRange(`C${7+i}`).values=[[x]]);
for (let c=0;c<5;c++) {
  const col = ["D","E","F","G","H"][c];
  const stCol = ["D","E","F","G","H"][c];
  const prev = c === 0 ? stCol : ["D","E","F","G"][c-1];
  const r = Object.fromEntries(ratioNames.map((x,i)=>[x,7+i]));
  ratios.getRange(`${col}${r["Revenue Growth"]}`).formulas=[[`='Statements'!${stCol}10`]];
  ratios.getRange(`${col}${r["Gross Margin"]}`).formulas=[[`='Statements'!${stCol}13`]];
  ratios.getRange(`${col}${r["Operating Margin"]}`).formulas=[[`='Statements'!${stCol}19`]];
  ratios.getRange(`${col}${r["Net Margin"]}`).formulas=[[`='Statements'!${stCol}25`]];
  ratios.getRange(`${col}${r["SG&A as % Revenue"]}`).formulas=[[`='Statements'!${stCol}16`]];
  ratios.getRange(`${col}${r["Current Ratio"]}`).formulas=[[`='Statements'!${stCol}39/'Statements'!${stCol}44`]];
  ratios.getRange(`${col}${r["Quick Ratio"]}`).formulas=[[`=('Statements'!${stCol}36+'Statements'!${stCol}37)/'Statements'!${stCol}44`]];
  ratios.getRange(`${col}${r["Debt / Equity"]}`).formulas=[[`=('Statements'!${stCol}43+'Statements'!${stCol}45)/'Statements'!${stCol}48`]];
  ratios.getRange(`${col}${r["Debt / Assets"]}`).formulas=[[`=('Statements'!${stCol}43+'Statements'!${stCol}45)/'Statements'!${stCol}40`]];
  ratios.getRange(`${col}${r["CFO / Net Income"]}`).formulas=[[`='Statements'!${stCol}30`]];
  ratios.getRange(`${col}${r["Free Cash Flow Margin"]}`).formulas=[[`='Statements'!${stCol}32/'Statements'!${stCol}9`]];
  ratios.getRange(`${col}${r["Inventory Turnover"]}`).formulas=[[`='Statements'!${stCol}11/(('Statements'!${stCol}38+'Statements'!${prev}38)/2)`]];
  ratios.getRange(`${col}${r["Receivable Days"]}`).formulas=[[`=(('Statements'!${stCol}37+'Statements'!${prev}37)/2)/'Statements'!${stCol}9*365`]];
  ratios.getRange(`${col}${r["Inventory Days"]}`).formulas=[[`=(('Statements'!${stCol}38+'Statements'!${prev}38)/2)/'Statements'!${stCol}11*365`]];
  ratios.getRange(`${col}${r["Payable Days"]}`).formulas=[[`=(('Statements'!${stCol}42+'Statements'!${prev}42)/2)/'Statements'!${stCol}11*365`]];
  ratios.getRange(`${col}${r["Cash Conversion Cycle"]}`).formulas=[[`=${col}${r["Receivable Days"]}+${col}${r["Inventory Days"]}-${col}${r["Payable Days"]}`]];
}
ratios.getRange("D7:H17").format.numberFormat = pctFmt;
for (const row of [12,13,14,15,16,18]) ratios.getRange(`D${row}:H${row}`).format.numberFormat = ratioFmt;
ratios.getRange("D19:H22").format.numberFormat = '0.0 "days"';
ratios.getRange("C24:H24").values = [["Working-capital balance changes ($M)", 2021, 2022, 2023, 2024, 2025]];
header(ratios.getRange("C24:H24"));
ratios.getRange("C25:C28").values = [["Accounts Receivable"],["Inventory"],["Accounts Payable"],["Net operating working capital"]];
for(let c=0;c<5;c++){
  const col=["D","E","F","G","H"][c], st=["D","E","F","G","H"][c];
  ratios.getRange(`${col}25`).formulas=[[`='Statements'!${st}37`]];
  ratios.getRange(`${col}26`).formulas=[[`='Statements'!${st}38`]];
  ratios.getRange(`${col}27`).formulas=[[`='Statements'!${st}42`]];
  ratios.getRange(`${col}28`).formulas=[[`=${col}25+${col}26-${col}27`]];
}
ratios.getRange("D25:H28").format.numberFormat = moneyFmt;
ratios.getRange("C28:H28").format.font = { name: font, size: 10, bold: true, color: dark };
setWidths(ratios);
ratios.getRange("D:H").format.columnWidth = 14;

// Graduate-level return, earnings-quality, driver, and acquisition analysis.
title(advanced, "Advanced performance analysis", "Returns, earnings quality, operating-profit attribution, and acquisition-adjusted growth.");
advanced.getRange("C6:H6").values = [["Metric",2021,2022,2023,2024,2025]];
header(advanced.getRange("C6:H6"));
const advancedMetrics = ["NOPAT","Average Assets","ROA","Invested Capital","Average Invested Capital","ROIC","Asset Turnover","Interest Coverage","Accrual Ratio","FCF / Net Income","Dividend Coverage","Capex / D&A","Net Operating Working Capital","NOWC / Revenue"];
advancedMetrics.forEach((x,i)=>advanced.getRange(`C${7+i}`).values=[[x]]);
const ar = Object.fromEntries(advancedMetrics.map((x,i)=>[x,7+i]));
for(let c=0;c<5;c++){
  const col=["D","E","F","G","H"][c], st=["D","E","F","G","H"][c], prev=c===0?st:["D","E","F","G"][c-1];
  advanced.getRange(`${col}${ar["NOPAT"]}`).formulas=[[`='Statements'!${st}18*(1-'Statements'!${st}23/'Statements'!${st}22)`]];
  advanced.getRange(`${col}${ar["Average Assets"]}`).formulas=[[`=AVERAGE('Statements'!${prev}40,'Statements'!${st}40)`]];
  advanced.getRange(`${col}${ar["ROA"]}`).formulas=[[`='Statements'!${st}24/${col}${ar["Average Assets"]}`]];
  advanced.getRange(`${col}${ar["Invested Capital"]}`).formulas=[[`='Statements'!${st}43+'Statements'!${st}45+'Statements'!${st}48-'Statements'!${st}36`]];
  advanced.getRange(`${col}${ar["Average Invested Capital"]}`).formulas=[[c===0?`=${col}${ar["Invested Capital"]}`:`=AVERAGE(${prev}${ar["Invested Capital"]},${col}${ar["Invested Capital"]})`]];
  advanced.getRange(`${col}${ar["ROIC"]}`).formulas=[[`=${col}${ar["NOPAT"]}/${col}${ar["Average Invested Capital"]}`]];
  advanced.getRange(`${col}${ar["Asset Turnover"]}`).formulas=[[`='Statements'!${st}9/${col}${ar["Average Assets"]}`]];
  advanced.getRange(`${col}${ar["Interest Coverage"]}`).formulas=[[`='Statements'!${st}18/'Statements'!${st}21`]];
  advanced.getRange(`${col}${ar["Accrual Ratio"]}`).formulas=[[`=('Statements'!${st}24-'Statements'!${st}29)/${col}${ar["Average Assets"]}`]];
  advanced.getRange(`${col}${ar["FCF / Net Income"]}`).formulas=[[`='Statements'!${st}32/'Statements'!${st}24`]];
  const dividendRow=dataRow["Dividends Paid"];
  advanced.getRange(`${col}${ar["Dividend Coverage"]}`).formulas=[[`='Statements'!${st}29/'Data'!${["M","N","O","P","Q"][c]}${dividendRow}`]];
  const daRow=dataRow["D&A Add-back"];
  advanced.getRange(`${col}${ar["Capex / D&A"]}`).formulas=[[`='Statements'!${st}31/'Data'!${["M","N","O","P","Q"][c]}${daRow}`]];
  advanced.getRange(`${col}${ar["Net Operating Working Capital"]}`).formulas=[[`='Statements'!${st}37+'Statements'!${st}38-'Statements'!${st}42`]];
  advanced.getRange(`${col}${ar["NOWC / Revenue"]}`).formulas=[[`=${col}${ar["Net Operating Working Capital"]}/'Statements'!${st}9`]];
}
advanced.getRange("D7:H8").format.numberFormat = moneyFmt;
advanced.getRange("D10:H11").format.numberFormat = moneyFmt;
advanced.getRange("D19:H19").format.numberFormat = moneyFmt;
for(const row of [ar["ROA"],ar["ROIC"],ar["Accrual Ratio"],ar["NOWC / Revenue"]]) advanced.getRange(`D${row}:H${row}`).format.numberFormat=pctFmt;
for(const row of [ar["Asset Turnover"],ar["Interest Coverage"],ar["FCF / Net Income"],ar["Dividend Coverage"],ar["Capex / D&A"]]) advanced.getRange(`D${row}:H${row}`).format.numberFormat=ratioFmt;

advanced.getRange("C23:G23").values = [["Operating-income attribution ($M)",2022,2023,2024,2025]];
header(advanced.getRange("C23:G23"));
const bridgeLabels=["Revenue effect at prior-year gross margin","Gross-margin rate effect","SG&A change","D&A change","Explained operating-income change","Reported operating-income change","Unexplained difference"];
bridgeLabels.forEach((x,i)=>advanced.getRange(`C${24+i}`).values=[[x]]);
for(let c=0;c<4;c++){
  const col=["D","E","F","G"][c], cur=["E","F","G","H"][c], prev=["D","E","F","G"][c];
  advanced.getRange(`${col}24`).formulas=[[`=('Statements'!${cur}9-'Statements'!${prev}9)*'Statements'!${prev}13`]];
  advanced.getRange(`${col}25`).formulas=[[`='Statements'!${cur}9*('Statements'!${cur}13-'Statements'!${prev}13)`]];
  advanced.getRange(`${col}26`).formulas=[[`=-('Statements'!${cur}15-'Statements'!${prev}15)`]];
  advanced.getRange(`${col}27`).formulas=[[`=-('Statements'!${cur}17-'Statements'!${prev}17)`]];
  advanced.getRange(`${col}28`).formulas=[[`=SUM(${col}24:${col}27)`]];
  advanced.getRange(`${col}29`).formulas=[[`='Statements'!${cur}18-'Statements'!${prev}18`]];
  advanced.getRange(`${col}30`).formulas=[[`=${col}28-${col}29`]];
}
advanced.getRange("D24:G30").format.numberFormat=moneyFmt;
advanced.getRange("C28:G30").format.font={name:font,size:10,bold:true,color:dark};

advanced.getRange("C33:F33").values = [["Acquisition-adjusted revenue ($M)",2023,2024,2025]];
header(advanced.getRange("C33:F33"));
advanced.getRange("C34:C39").values = [["Primary segment revenue"],["Other segment revenue"],["Total revenue"],["Primary segment growth"],["Other segment contribution to annual growth"],["Primary segment contribution to annual growth"]];
advanced.getRange("D34:F36").values = [[152669,153108,151966],[0,6406,12717],[152669,159514,164683]];
advanced.getRange("D37").formulas=[[`=""`]];
advanced.getRange("E37:F37").formulas=[["=E34/D34-1","=F34/E34-1"]];
advanced.getRange("D38").formulas=[[`=""`]];
advanced.getRange("E38:F38").formulas=[["=E35-D35","=F35-E35"]];
advanced.getRange("D39").formulas=[[`=""`]];
advanced.getRange("E39:F39").formulas=[["=E34-D34","=F34-E34"]];
advanced.getRange("D34:F36").format.numberFormat=moneyFmt;
advanced.getRange("D37:F37").format.numberFormat=pctFmt;
advanced.getRange("D38:F39").format.numberFormat=moneyFmt;
advanced.getRange("H33:K33").values = [["Product category revenue ($M)",2023,2024,2025]];
header(advanced.getRange("H33:K33"));
advanced.getRange("H34:H36").values = [["Building Materials"],["Décor"],["Hardlines"]];
advanced.getRange("I34:K36").values = [[52572,52862,52439],[52750,52525,51679],[47347,47721,47848]];
advanced.getRange("I34:K36").format.numberFormat=moneyFmt;
advanced.getRange("C42:K42").values=[["Interpretive finding","Evidence","Implication","Decision threshold",null,null,null,null,null]];
header(advanced.getRange("C42:K42"));
advanced.getRange("C43:F47").values=[
  ["Revenue quality","FY2025 total revenue rose while primary segment revenue declined","Acquired businesses were the main source of reported growth","Require positive primary-segment growth before assuming broad demand recovery"],
  ["Return efficiency","ROIC declined as invested capital expanded","Acquisitions and margin compression reduced capital productivity","Monitor ROIC recovery relative to acquisition integration"],
  ["Earnings quality","CFO remained above net income but conversion declined","Earnings remained cash-backed, with weaker working-capital support","CFO / net income below 1.0x would be a warning"],
  ["Profit bridge","Higher SG&A offset most gross-profit growth","Cost absorption is the central margin issue","Base thesis requires operating margin at or above 13%"],
  ["Liquidity","Inventory remains the dominant current asset","Current ratio alone overstates immediate liquidity","Track inventory days and payable days together"]
];
advanced.getRange("C43:F47").format.wrapText=true;
advanced.getRange("C43:F47").format.verticalAlignment="top";
advanced.getRange("43:47").format.rowHeight=42;
setWidths(advanced);
advanced.getRange("C:C").format.columnWidth=38;
advanced.getRange("D:K").format.columnWidth=17;
advanced.freezePanes.freezeRows(6);
advanced.getRange("D:F").format.columnWidth=27;
advanced.getRange("G:G").format.columnWidth=17;
advanced.getRange("H:H").format.columnWidth=28;
advanced.getRange("I:K").format.columnWidth=17;

// Scenario comparison from the reproducible Python output; clearly labeled refresh behavior.
title(scenario, "Scenario outlook", "Scenario outputs are refreshed by running scripts/transform_financials.py. The active Statements forecast uses the selector on Assumptions.");
scenario.getRange("C6:F6").values = [["Metric", "Downside", "Base", "Upside"]];
header(scenario.getRange("C6:F6"));
const scenarioMetrics = ["Revenue", "Operating Income", "Net Income", "Operating Cash Flow"];
let sr = 7;
for (const year of [2026,2027,2028]) {
  section(scenario, `C${sr}:F${sr}`, `${year} outlook`);
  sr++;
  for (const metric of scenarioMetrics) {
    scenario.getRange(`C${sr}`).values = [[metric]];
    const vals = ["Downside","Base","Upside"].map(sc => model.scenarios.find(x => x.Scenario===sc && x["Fiscal Year"]===year && x.Metric===metric)["Value USD millions"]);
    scenario.getRange(`D${sr}:F${sr}`).values = [vals];
    sr++;
  }
  sr++;
}
scenario.getRange("D7:F30").format.numberFormat = moneyFmt;
scenario.getRange("E7:E30").format.fill = paleBlue;
scenario.getRange("C29:F29").values = [["Interpretation", "Downside", "Base", "Upside"]];
header(scenario.getRange("C29:F29"));
scenario.getRange("C30:F33").values = [
  ["Demand", "Flat near-term sales", "Low-single-digit growth", "Pro demand and housing recovery"],
  ["Margins", "Limited gross-margin recovery", "Gradual operating leverage", "Faster mix and productivity gains"],
  ["Cash", "CFO conversion stays near 1.1x", "Healthy conversion funds capex/dividends", "Stronger earnings and conversion"],
  ["Key risk", "High rates and weak big-ticket demand", "Execution on acquired businesses", "Upside requires macro improvement"],
];
scenario.getRange("C30:F33").format.wrapText = true;
setWidths(scenario);
scenario.getRange("D:F").format.columnWidth = 24;
scenario.getRange("C:C").format.columnWidth = 30;

// Formula-based two-driver sensitivity for FY2028 operating income.
title(sensitivity, "Sensitivity analysis", "FY2028 operating income under alternative revenue CAGR and operating-margin assumptions.");
sensitivity.getRange("C6:D9").values=[["Reference input","Value"],["FY2025 Revenue",null],["Base FY2028 Revenue",null],["Base FY2028 Operating Income",null]];
header(sensitivity.getRange("C6:D6"));
sensitivity.getRange("D7:D9").formulas=[["='Statements'!H9"],["='Statements'!K9"],["='Statements'!K18"]];
sensitivity.getRange("D7:D9").format.numberFormat=moneyFmt;
sensitivity.getRange("C12:J12").values=[["3-year Revenue CAGR / Operating Margin",0.10,0.11,0.12,0.13,0.14,0.15,0.16]];
header(sensitivity.getRange("C12:J12"));
const growthRates=[0,0.01,0.02,0.03,0.04,0.05,0.06];
for(let i=0;i<growthRates.length;i++){
  const row=13+i;
  sensitivity.getRange(`C${row}`).values=[[growthRates[i]]];
  for(let j=0;j<7;j++){
    const col=["D","E","F","G","H","I","J"][j];
    const marginCell=["D","E","F","G","H","I","J"][j];
    sensitivity.getRange(`${col}${row}`).formulas=[[`=$D$7*(1+$C${row})^3*${marginCell}$12`]];
  }
}
sensitivity.getRange("C13:C19").format.numberFormat=pctFmt;
sensitivity.getRange("D12:J12").format.numberFormat=pctFmt;
sensitivity.getRange("D13:J19").format.numberFormat=moneyFmt;
sensitivity.getRange("D13:J19").conditionalFormats.add("colorScale",{colors:["#F4CCCC","#FFF2CC","#D9EAD3"],thresholds:["min",{type:"percentile",value:50},"max"]});
sensitivity.getRange("C22:F22").values=[["Decision framework","Trigger","Interpretation","Response"]];
header(sensitivity.getRange("C22:F22"));
sensitivity.getRange("C23:F26").values=[
  ["Upside confirmation","Primary-segment growth turns positive and operating margin exceeds 13%","Demand and integration improve together","Revisit the upside case"],
  ["Base case intact","Revenue grows 2%–4%, margin reaches 12.8%–13.2%, CFO conversion exceeds 1.1x","Cash generation supports capex and dividends","Maintain neutral stance"],
  ["Downside warning","CFO conversion falls below 1.0x or inventory days rise with flat sales","Working capital weakens earnings quality","Reduce forecast assumptions"],
  ["Capital-efficiency warning","ROIC continues falling despite acquisition integration","Incremental invested capital is not earning an adequate return","Require a lower reinvestment case"]
];
sensitivity.getRange("C23:F26").format.wrapText=true;
sensitivity.getRange("23:26").format.rowHeight=44;
setWidths(sensitivity);
sensitivity.getRange("C:C").format.columnWidth=34;
sensitivity.getRange("D:J").format.columnWidth=18;
sensitivity.getRange("D:D").format.columnWidth=30;
sensitivity.getRange("E:E").format.columnWidth=27;
sensitivity.getRange("F:F").format.columnWidth=23;

// Dashboard.
title(dashboard, "Home Depot financial performance", "Fiscal 2021-2025 actuals with a three-year scenario outlook. USD millions unless noted.");
dashboard.getRange("C6:D6").values = [["Case Selected:", null]];
dashboard.getRange("D6").formulas = [["='Statements'!D6"]];
dashboard.getRange("D6").format.font = { name: font, size: 11, bold: true, color: green };
dashboard.getRange("C8:E8").values = [["FY2025 Revenue", "FY2025 Operating Margin", "FY2025 CFO / Net Income"]];
header(dashboard.getRange("C8:E8"));
dashboard.getRange("C9:E9").formulas = [["='Statements'!H9", "='Statements'!H19", "='Statements'!H30"]];
dashboard.getRange("C9").format.numberFormat = moneyFmt;
dashboard.getRange("D9").format.numberFormat = pctFmt;
dashboard.getRange("E9").format.numberFormat = ratioFmt;
dashboard.getRange("C9:E9").format.font = { name: font, size: 14, bold: true, color: navy };
dashboard.getRange("G8:I8").values = [["FY2025 Current Ratio", "FY2025 ROIC", "FY2025 Free Cash Flow"]];
header(dashboard.getRange("G8:I8"));
dashboard.getRange("G9:I9").formulas = [["='Ratios & WC'!H12", "='Advanced Analysis'!H12", "='Statements'!H32"]];
dashboard.getRange("G9").format.numberFormat = ratioFmt;
dashboard.getRange("H9").format.numberFormat = pctFmt;
dashboard.getRange("I9").format.numberFormat = moneyFmt;
dashboard.getRange("G9:I9").format.font = { name: font, size: 14, bold: true, color: navy };

dashboard.getRange("C12:F12").values = [["Executive view", "Evidence", "Assessment", "Action"]];
header(dashboard.getRange("C12:F12"));
dashboard.getRange("C13:F17").values = [
  ["Growth", "Revenue reached $164.7B in FY2025", "Acquisitions offset soft organic demand", "Separate acquired growth from same-store recovery"],
  ["Profitability", "Operating margin fell from 15.2% to 12.7% since FY2021", "SG&A and mix diluted operating leverage", "Prioritize integration savings and gross-margin discipline"],
  ["Cash flow", "CFO exceeded net income in four of five years", "Cash generation remains solid but weakened in FY2025", "Track inventory and payables against demand"],
  ["Liquidity", "Current ratio remained near 1.0x", "Adequate but limited short-term cushion", "Protect cash conversion during slower demand"],
  ["Recommendation", "Base case supports stable cash generation", "Hold / monitor execution", "Reassess if margins stabilize above 13% or demand weakens"],
];
dashboard.getRange("C13:F17").format.wrapText = true;
dashboard.getRange("C13:F17").format.verticalAlignment = "top";
dashboard.getRange("C13:C17").format.font = { name: font, size: 10, bold: true, color: navy };
dashboard.getRange("C20:K20").values = [["Fiscal Year",2021,2022,2023,2024,2025,2026,2027,2028]];
header(dashboard.getRange("C20:K20"));
dashboard.getRange("C21:C24").values = [["Revenue"],["Operating Income"],["Operating Margin"],["CFO / Net Income"]];
dashboard.getRange("D21:K24").formulas = [
  cols.map(c=>`='Statements'!${c}9`), cols.map(c=>`='Statements'!${c}18`), cols.map(c=>`='Statements'!${c}19`), cols.map(c=>`='Statements'!${c}30`)
];
dashboard.getRange("D21:K22").format.numberFormat = moneyFmt;
dashboard.getRange("D23:K23").format.numberFormat = pctFmt;
dashboard.getRange("D24:K24").format.numberFormat = ratioFmt;

dashboard.getRange("Q20:S20").values = [["Fiscal Year", "Revenue", "Operating Income"]];
for (let i=0;i<8;i++) {
  const row=21+i;
  const src=cols[i];
  dashboard.getRange(`Q${row}:S${row}`).formulas = [[`=${src}20`, `=${src}21`, `=${src}22`]];
}
dashboard.getRange("R21:S28").format.numberFormat = moneyFmt;
const trendChart = dashboard.charts.add("line", dashboard.getRange("Q20:S28"));
trendChart.title = "Revenue and Operating Income Trend ($M)";
trendChart.titleTextStyle.typeface = font;
trendChart.titleTextStyle.fontSize = 12;
trendChart.legend = { position: "top", textStyle: { typeface: font } };
trendChart.xAxis = { axisType: "textAxis", textStyle: { typeface: font, fontSize: 10 } };
trendChart.yAxis = { numberFormatCode: '$0,"B"', numberFormatSourceLinked: false, textStyle: { typeface: font } };
trendChart.setPosition("C26", "I42");
if (trendChart.series.items[0]) trendChart.series.items[0].line = { fill: blue, style: "solid", width: 2 };
if (trendChart.series.items[1]) trendChart.series.items[1].line = { fill: green, style: "solid", width: 2 };
dashboard.getRange("M20:O20").values = [["Fiscal Year", "Operating Margin", "CFO / Net Income"]];
for (let i=0;i<8;i++) {
  const row=21+i;
  const src=cols[i];
  dashboard.getRange(`M${row}:O${row}`).formulas = [[`=${src}20`, `=${src}23`, `=${src}24`]];
}
dashboard.getRange("N21:N28").format.numberFormat = pctFmt;
dashboard.getRange("O21:O28").format.numberFormat = ratioFmt;
const marginChart = dashboard.charts.add("line", dashboard.getRange("M20:O28"));
marginChart.title = "Margin and Cash Conversion";
marginChart.titleTextStyle.typeface = font;
marginChart.titleTextStyle.fontSize = 12;
marginChart.legend = { position: "top", textStyle: { typeface: font } };
marginChart.xAxis = { axisType: "textAxis", textStyle: { typeface: font, fontSize: 10 } };
marginChart.yAxis = { numberFormatCode: "0.0x", numberFormatSourceLinked: false, textStyle: { typeface: font } };
marginChart.setPosition("J26", "P42");
setWidths(dashboard);
dashboard.getRange("C:F").format.columnWidth = 26;
dashboard.getRange("G:I").format.columnWidth = 19;
dashboard.getRange("J:P").format.columnWidth = 12;
dashboard.getRange("13:17").format.rowHeight = 42;

// Audit checks.
title(audit, "Audit checks", "Differences are shown in USD millions. Values above 0.01 require review.");
audit.getRange("C6:H6").values = [["Check",2021,2022,2023,2024,2025]];
header(audit.getRange("C6:H6"));
audit.getRange("C7:C11").values = [["Gross profit: revenue less cost of sales"],["Balance sheet: assets less liabilities and equity"],["Operating expenses: gross profit less operating income vs SG&A plus D&A"],["Operating-income attribution: explained less reported change"],["Acquisition-adjusted revenue: Primary plus Other less Total"]];
for(let c=0;c<5;c++){
  const col=["D","E","F","G","H"][c], st=["D","E","F","G","H"][c];
  audit.getRange(`${col}7`).formulas=[[`='Statements'!${st}12-('Statements'!${st}9-'Statements'!${st}11)`]];
  audit.getRange(`${col}8`).formulas=[[`='Statements'!${st}40-'Statements'!${st}47-'Statements'!${st}48`]];
  audit.getRange(`${col}9`).formulas=[[`=('Statements'!${st}12-'Statements'!${st}18)-('Statements'!${st}15+'Statements'!${st}17)`]];
}
audit.getRange("D10:H10").formulas=[[`=""`,`='Advanced Analysis'!D30`,`='Advanced Analysis'!E30`,`='Advanced Analysis'!F30`,`='Advanced Analysis'!G30`]];
audit.getRange("D11:H11").formulas=[[`=""`,`=""`,`='Advanced Analysis'!D34+'Advanced Analysis'!D35-'Advanced Analysis'!D36`,`='Advanced Analysis'!E34+'Advanced Analysis'!E35-'Advanced Analysis'!E36`,`='Advanced Analysis'!F34+'Advanced Analysis'!F35-'Advanced Analysis'!F36`]];
audit.getRange("D7:H11").format.numberFormat = '0.00;[Red](0.00);0.00';
audit.getRange("D7:H11").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0.01, format: { fill: "#FCE4D6", font: { bold: true, color: red } } });
audit.getRange("D7:H11").conditionalFormats.add("cellIs", { operator: "lessThan", formula: -0.01, format: { fill: "#FCE4D6", font: { bold: true, color: red } } });
setWidths(audit);
audit.getRange("D:H").format.columnWidth = 14;

wb.recalculate();
await fs.mkdir(outputDir, { recursive: true });
const checks = await wb.inspect({ kind: "table", range: "Dashboard!C6:K24", include: "values,formulas", tableMaxRows: 25, tableMaxCols: 12 });
console.log(checks.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" });
console.log(errors.ndjson);
for (const sheetName of ["Dashboard","Assumptions","Statements","Ratios & WC","Advanced Analysis","Scenario Outlook","Sensitivity","Data","Audit"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName.replace(/[^A-Za-z0-9]+/g,"_")}.png`), new Uint8Array(await preview.arrayBuffer()));
}
const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(path.join(outputDir, "Home_Depot_Financial_Analysis.xlsx"));
console.log(path.join(outputDir, "Home_Depot_Financial_Analysis.xlsx"));
