#requires -Version 5.1
<#
PowerBI Analytics Setup - on premises lab
Run on a Windows computer with network access to SQL Server, as an account permitted to create views in PowerBIAnalytics_PROD.
No database recreation, no data deletion, no modifications to Agent jobs.
Example: .\04_PowerBI_Analytics_Setup.ps1 -Server '192.168.88.18' -Database 'PowerBIAnalytics_PROD' -OutputFolder 'C:\SQLLab\PowerBI'
Optional: -OpenPowerBI (opens PBIDesktop.exe if installed; it cannot automatically assemble a PBIX)
#>
[CmdletBinding()]
param(
 [string]$Server = '192.168.88.18',
 [string]$Database = 'PowerBIAnalytics_PROD',
 [string]$OutputFolder = 'C:\SQLLab\PowerBI',
 [switch]$OpenPowerBI
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($Database -notmatch '^PowerBIAnalytics_(PROD|DEV)$') { throw 'Only the synthetic PowerBIAnalytics_PROD or PowerBIAnalytics_DEV lab databases are allowed.' }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$cs = "Server=$Server;Database=$Database;Integrated Security=SSPI;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15;"
$conn = New-Object System.Data.SqlClient.SqlConnection($cs)
function Invoke-SqlNonQuery([string]$Sql) {
 $cmd=$conn.CreateCommand(); $cmd.CommandTimeout=120; $cmd.CommandText=$Sql
 try {[void]$cmd.ExecuteNonQuery()} finally {$cmd.Dispose()}
}
function Get-SqlTable([string]$Sql) {
 $cmd=$conn.CreateCommand(); $cmd.CommandTimeout=120; $cmd.CommandText=$Sql
 $da=New-Object System.Data.SqlClient.SqlDataAdapter($cmd); $dt=New-Object System.Data.DataTable
 try {[void]$da.Fill($dt); return ,$dt} finally {$cmd.Dispose();$da.Dispose()}
}
try {
 $conn.Open()
 Write-Host "Connected: $Server / $Database" -ForegroundColor Green
 $required = @('Customers','Products','SalesOrders','SalesOrderDetails','Payments','SalesTargets')
 $found = Get-SqlTable "SELECT name FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo')"
 foreach($name in $required) { if(-not (@($found.Rows | ForEach-Object { $_['name'].ToString() }) -contains $name)) { throw "Missing dbo.$name. Schema must be inspected before constructing a reporting view." } }
 # Get actual column metadata; do not assume schema created by an earlier script.
 $columns=Get-SqlTable "SELECT t.name AS TableName,c.name AS ColumnName FROM sys.tables t JOIN sys.columns c ON t.object_id=c.object_id WHERE t.schema_id=SCHEMA_ID('dbo') ORDER BY t.name,c.column_id"
 $columns | Export-Csv (Join-Path $OutputFolder 'SQL_Column_Inventory.csv') -NoTypeInformation -Encoding UTF8
 $existing=Get-SqlTable "SELECT name FROM sys.views WHERE schema_id=SCHEMA_ID('dbo')"
 if(-not (@($existing.Rows | ForEach-Object { $_['name'].ToString() }) -contains 'vw_SalesAnalytics')) { throw 'Missing dbo.vw_SalesAnalytics. Run the original PROD setup script successfully before this script.' }
 $viewColumns=Get-SqlTable "SELECT c.name AS ColumnName,ty.name AS DataType FROM sys.columns c JOIN sys.views v ON c.object_id=v.object_id JOIN sys.types ty ON c.user_type_id=ty.user_type_id WHERE v.object_id=OBJECT_ID('dbo.vw_SalesAnalytics') ORDER BY c.column_id"
 $viewColumns | Export-Csv (Join-Path $OutputFolder 'SalesAnalytics_Columns.csv') -NoTypeInformation -Encoding UTF8
 $check=Get-SqlTable "SELECT COUNT_BIG(*) AS SalesAnalyticsRows FROM dbo.vw_SalesAnalytics"
 $check | Export-Csv (Join-Path $OutputFolder 'SalesAnalytics_RowCount.csv') -NoTypeInformation -Encoding UTF8
 # Output a Power Query M query that uses the established reporting view; no passwords included.
 $escapedServer=$Server.Replace('"','""'); $escapedDb=$Database.Replace('"','""')
 $m=@"
let
    Source = Sql.Database("$escapedServer", "$escapedDb", [CreateNavigationProperties=false]),
    SalesAnalytics = Source{[Schema="dbo",Item="vw_SalesAnalytics"]}[Data]
in
    SalesAnalytics
"@
 Set-Content -Path (Join-Path $OutputFolder 'SalesAnalytics_PowerQuery.m') -Value $m -Encoding UTF8
 $instructions=@"
POWER BI DESKTOP (FREE) - LOCAL REPORT INSTRUCTIONS
1. Open Power BI Desktop on a Windows computer with connectivity to $Server.
2. Home > Get Data > Blank Query > Advanced Editor.
3. Paste contents of SalesAnalytics_PowerQuery.m; click Done.
4. Choose Windows authentication with a dedicated read-only SQL user/group (do not share service-account credentials).
5. Close & Apply; save as PowerBIAnalytics_Report.pbix.
6. In Data pane, check columns using SalesAnalytics_Columns.csv and choose Date/Decimal types where appropriate.
7. Create visuals: a table with sales rows, bar chart by category/product if such columns exist, and date trend if a sales date column exists.
8. Use Home > Refresh to refresh imported data. Free Desktop supports local report authoring; automatic unattended PBIX refresh needs an appropriate platform/gateway/licensing solution.
IMPORTANT: This script does NOT fabricate a PBIX or change PROD tables. All business records remain synthetic lab data.
"@
 Set-Content -Path (Join-Path $OutputFolder 'README_PowerBI.txt') -Value $instructions -Encoding UTF8
 Write-Host "View rows: $($check.Rows[0]['SalesAnalyticsRows'])" -ForegroundColor Green
 Write-Host "Created Power BI M query, schema inventory, and instructions under $OutputFolder" -ForegroundColor Green
 $viewColumns | Format-Table -AutoSize
} finally { $conn.Dispose() }
if($OpenPowerBI) {
 $app=Get-Command PBIDesktop.exe -ErrorAction SilentlyContinue
 if($app){ Start-Process -FilePath $app.Source } else { Write-Warning 'PBIDesktop.exe not found on PATH. Open Power BI Desktop manually.' }
}
