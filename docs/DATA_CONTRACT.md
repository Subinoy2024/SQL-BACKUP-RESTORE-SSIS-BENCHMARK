# Input and destination data contract

The benchmark source is an `.xlsx` workbook with a worksheet named **`SalesData$`** (the trailing `$` is part of the ACE sheet identifier) and exactly these expected input field names:

| Excel field | Expected meaning | SQL destination type |
|---|---|---|
| OrderID | integer order identifier | INT |
| ProductID | integer product identifier | INT |
| Quantity | integer quantity | INT |
| SalesAmount | numeric sales amount | FLOAT |

The SQL destination is `PowerBIAnalytics_DEV.dbo.SSIS_Benchmark_Results` with identity `BenchmarkID` and default `LoadTimestamp` that are **not** mapped from Excel. Create it via `sql/01_Benchmark_Table.sql`.

ACE infers Excel types from cells. Keep the entire source column consistent: use numeric cells (not text such as `1,000` or `$123`); avoid blank headers, errors, formulas with mixed return types, and merged cells. Confirm that all 100,000 source records have the required columns. For a new schema, update the typed C# mapping names and table DDL together and revalidate before loading.

The original input workbook is not included to keep the repository small and free from binary data. Create a sanitized, deterministic dataset within your own lab.
