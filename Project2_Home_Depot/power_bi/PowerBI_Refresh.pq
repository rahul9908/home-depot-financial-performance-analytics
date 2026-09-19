// Create three blank queries named Financials, Ratios, and Scenarios.
// Replace ProjectFolder below with the local project path, then paste the
// matching expression into each query's Advanced Editor.

// Financials
let
    ProjectFolder = "C:\\Users\\rahul\\Documents\\ChatGPT\\FA2\\Project2_Home_Depot\\data\\",
    Source = Csv.Document(File.Contents(ProjectFolder & "financials_long.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Types = Table.TransformColumnTypes(Headers, {{"Company", type text}, {"Ticker", type text}, {"Fiscal Year", Int64.Type}, {"Fiscal Year End", type date}, {"Statement", type text}, {"Line Item", type text}, {"Value USD millions", type number}, {"Source", type text}})
in
    Types

// Ratios
let
    ProjectFolder = "C:\\Users\\rahul\\Documents\\ChatGPT\\FA2\\Project2_Home_Depot\\data\\",
    Source = Csv.Document(File.Contents(ProjectFolder & "ratios_long.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Types = Table.TransformColumnTypes(Headers, {{"Company", type text}, {"Ticker", type text}, {"Fiscal Year", Int64.Type}, {"Metric", type text}, {"Value", type number}})
in
    Types

// Scenarios
let
    ProjectFolder = "C:\\Users\\rahul\\Documents\\ChatGPT\\FA2\\Project2_Home_Depot\\data\\",
    Source = Csv.Document(File.Contents(ProjectFolder & "scenario_outlook.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Types = Table.TransformColumnTypes(Headers, {{"Scenario", type text}, {"Fiscal Year", Int64.Type}, {"Metric", type text}, {"Value USD millions", type number}})
in
    Types
