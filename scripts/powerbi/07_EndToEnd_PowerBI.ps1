#requires -Version 5.1
<#!
.SYNOPSIS
  Creates a new standalone Power BI Project for the PowerBIAnalytics_PROD lab.
.DESCRIPTION
  Validates SQL data, creates TMDL semantic model, PBIR report with visuals,
  plus deployment readme and diagnostics. Never changes SQL data, SQL Agent jobs,
  or the older AnalyticsLab project. Opening in Power BI Desktop is required
  to import data and save a PBIX.
.EXAMPLE
  .\07_EndToEnd_PowerBI.ps1 -SqlServer 192.168.88.18 -OutputRoot C:\SQLLab\PBI_Automated
#>
[CmdletBinding()]
param(
  [string]$SqlServer='192.168.88.18',
  [string]$Database='PowerBIAnalytics_PROD',
  [string]$OutputRoot='C:\SQLLab\PBI_Automated',
  [switch]$OpenWhenDone
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Write-JsonFile([string]$Path, [object]$Value) {
  $parent=Split-Path -Parent $Path
  if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
  $json=ConvertTo-Json -InputObject $Value -Depth 80
  [System.IO.File]::WriteAllText($Path,$json,[System.Text.UTF8Encoding]::new($false))
}
function Write-TextFile([string]$Path,[string]$Value) {
  $parent=Split-Path -Parent $Path
  if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
  [System.IO.File]::WriteAllText($Path,$Value,[System.Text.UTF8Encoding]::new($false))
}
function New-Identifier { [guid]::NewGuid().ToString('N').Substring(0,20) }

Write-Host '=== Checking PROD reporting view (READ ONLY) ===' -ForegroundColor Cyan
if ($SqlServer -notmatch '^[a-zA-Z0-9._-]+$' -or $Database -notmatch '^[a-zA-Z0-9_]+$') { throw 'Invalid server/database identifier.' }
$cs="Server=$SqlServer;Database=$Database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15"
$conn=[System.Data.SqlClient.SqlConnection]::new($cs)
try {
  $conn.Open()
  $cmd=$conn.CreateCommand()
  $cmd.CommandText="SELECT COUNT_BIG(*) FROM dbo.vw_SalesAnalytics"
  $cmd.CommandTimeout=60
  $rows=[long]$cmd.ExecuteScalar()
  $cmd.CommandText="SELECT TOP(1) OrderID,OrderDate,OrderStatus,CustomerID,CustomerName,Region,ProductID,ProductName,Category,Quantity,UnitPrice,SalesAmount FROM dbo.vw_SalesAnalytics"
  $reader=$cmd.ExecuteReader()
  $actual=@(); for ($i=0; $i -lt $reader.FieldCount; $i++) { $actual += $reader.GetName($i) }; $reader.Close()
} finally { $conn.Dispose() }
$required=@('OrderID','OrderDate','OrderStatus','CustomerID','CustomerName','Region','ProductID','ProductName','Category','Quantity','UnitPrice','SalesAmount')
foreach($name in $required){ if($actual -notcontains $name){throw "Missing SQL view column: $name"} }
Write-Host "Reporting rows: $rows; columns: $($actual.Count)" -ForegroundColor Green

# Do not overwrite an existing PBIP. Every run goes to a fresh folder.
New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
$run=Join-Path $OutputRoot ('SalesAnalytics_'+(Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $run -Force | Out-Null
$project=Join-Path $run 'SalesAnalytics.pbip'
$report=Join-Path $run 'SalesAnalytics.Report'
$model=Join-Path $run 'SalesAnalytics.SemanticModel'
$def=Join-Path $report 'definition'
$pages=Join-Path $def 'pages'
$pageId=New-Identifier
$page=Join-Path $pages $pageId

Write-JsonFile $project @{'$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/pbipProperties/1.0.0/schema.json';version='1.0';artifacts=@(@{report=@{path='SalesAnalytics.Report'}})}
Write-JsonFile (Join-Path $model 'definition.pbism') @{'$schema'='https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json';version='4.2';settings=@{qnaEnabled=$false}}
Write-TextFile (Join-Path $model 'definition\database.tmdl') "database SalesAnalytics`n`tcompatibilityLevel: 1702`n`tcompatibilityMode: powerBI`n"
Write-TextFile (Join-Path $model 'definition\model.tmdl') "model Model`n`tculture: en-US`n`tdefaultPowerBIDataSourceVersion: powerBI_V3`n`tsourceQueryCulture: en-US`n`nref table SalesAnalytics`n"
# TMDL tab indentations are significant. Query is Import using SQL view.
$escapedServer=$SqlServer.Replace('"','""')
$escapedDB=$Database.Replace('"','""')
$tmdl=@"
table SalesAnalytics
`tcolumn OrderID
`t`tdataType: int64
`t`tsourceColumn: OrderID
`t`tsummarizeBy: none

`tcolumn OrderDate
`t`tdataType: dateTime
`t`tsourceColumn: OrderDate
`t`tsummarizeBy: none

`tcolumn OrderStatus
`t`tdataType: string
`t`tsourceColumn: OrderStatus
`t`tsummarizeBy: none

`tcolumn CustomerID
`t`tdataType: int64
`t`tsourceColumn: CustomerID
`t`tsummarizeBy: none

`tcolumn CustomerName
`t`tdataType: string
`t`tsourceColumn: CustomerName
`t`tsummarizeBy: none

`tcolumn Region
`t`tdataType: string
`t`tsourceColumn: Region
`t`tsummarizeBy: none

`tcolumn ProductID
`t`tdataType: int64
`t`tsourceColumn: ProductID
`t`tsummarizeBy: none

`tcolumn ProductName
`t`tdataType: string
`t`tsourceColumn: ProductName
`t`tsummarizeBy: none

`tcolumn Category
`t`tdataType: string
`t`tsourceColumn: Category
`t`tsummarizeBy: none

`tcolumn Quantity
`t`tdataType: int64
`t`tsourceColumn: Quantity
`t`tsummarizeBy: sum

`tcolumn UnitPrice
`t`tdataType: decimal
`t`tsourceColumn: UnitPrice
`t`tsummarizeBy: none

`tcolumn SalesAmount
`t`tdataType: decimal
`t`tsourceColumn: SalesAmount
`t`tsummarizeBy: sum

`tmeasure 'Total Revenue' = SUM(SalesAnalytics[SalesAmount])
`t`tformatString: '#,##0.00'

`tmeasure 'Total Orders' = DISTINCTCOUNT(SalesAnalytics[OrderID])
`t`tformatString: '#,##0'

`tmeasure 'Customers Served' = DISTINCTCOUNT(SalesAnalytics[CustomerID])
`t`tformatString: '#,##0'

`tmeasure 'Total Quantity' = SUM(SalesAnalytics[Quantity])
`t`tformatString: '#,##0'

`tmeasure 'Average Order Value' = DIVIDE([Total Revenue], [Total Orders])
`t`tformatString: '#,##0.00'

`tpartition SalesAnalytics = m
`t`tmode: import
`t`tsource =
`t`t`tlet
`t`t`t    Source = Sql.Database("$escapedServer", "$escapedDB", [CreateNavigationProperties=false]),
`t`t`t    Sales = Source{[Schema="dbo",Item="vw_SalesAnalytics"]}[Data]
`t`t`tin
`t`t`t    Sales
"@
Write-TextFile (Join-Path $model 'definition\tables\SalesAnalytics.tmdl') $tmdl
Write-JsonFile (Join-Path $report 'definition.pbir') @{'$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json';version='4.0';datasetReference=@{byPath=@{path='../SalesAnalytics.SemanticModel'}}}
Write-JsonFile (Join-Path $def 'version.json') @{'$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json';version='2.0.0'}
Write-JsonFile (Join-Path $def 'report.json') @{'$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.0.0/schema.json';themeCollection=@{baseTheme=@{name='CY24SU07';reportVersionAtImport='5.54';type='SharedResources'}};settings=@{useStylableVisualContainerHeader=$true}}
Write-JsonFile (Join-Path $pages 'pages.json') @{'$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json';pageOrder=@($pageId);activePageName=$pageId}
Write-JsonFile (Join-Path $page 'page.json') @{'$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json';name=$pageId;displayName='Sales Analytics';displayOption='FitToPage';height=720;width=1280}

function New-ColumnProjection([string]$column,[string]$ref){
 return @{field=@{Column=@{Expression=@{SourceRef=@{Entity='SalesAnalytics'}};Property=$column}};queryRef=$ref;nativeQueryRef=$column}
}
function New-MeasureProjection([string]$measure){
 return @{field=@{Measure=@{Expression=@{SourceRef=@{Entity='SalesAnalytics'}};Property=$measure}};queryRef="SalesAnalytics.$measure";nativeQueryRef=$measure}
}
function New-AggregationProjection([string]$column){
 return @{field=@{Aggregation=@{Expression=@{Column=@{Expression=@{SourceRef=@{Entity='SalesAnalytics'}};Property=$column}};Function=0}};queryRef="Sum(SalesAnalytics.$column)";nativeQueryRef="Sum of $column"}
}
function Add-Visual([string]$type,[string]$title,[int]$x,[int]$y,[int]$w,[int]$h,[hashtable]$queryState){
 $id=New-Identifier
 $obj=@{
 '$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.7.0/schema.json'
 name=$id
 position=@{x=$x;y=$y;z=0;width=$w;height=$h;tabOrder=0}
 visual=@{visualType=$type;query=@{queryState=$queryState};drillFilterOtherVisuals=$true}
 }
 Write-JsonFile (Join-Path $page "visuals\$id\visual.json") $obj
 Write-Host "Added visual definition: $title ($type)"
}
# PBIR visual role names can change between Desktop releases; validate rendering after opening.
Add-Visual 'card' 'Total Revenue' 15 15 395 110 @{Values=@{projections=@((New-MeasureProjection 'Total Revenue'))}}
Add-Visual 'card' 'Total Orders' 425 15 395 110 @{Values=@{projections=@((New-MeasureProjection 'Total Orders'))}}
Add-Visual 'card' 'Customers Served' 835 15 395 110 @{Values=@{projections=@((New-MeasureProjection 'Customers Served'))}}
Add-Visual 'lineChart' 'Revenue by Date' 15 145 760 270 @{Category=@{projections=@((New-ColumnProjection 'OrderDate' 'SalesAnalytics.OrderDate'))};Y=@{projections=@((New-MeasureProjection 'Total Revenue'))}}
Add-Visual 'clusteredBarChart' 'Revenue by Region' 795 145 435 270 @{Category=@{projections=@((New-ColumnProjection 'Region' 'SalesAnalytics.Region'))};Y=@{projections=@((New-MeasureProjection 'Total Revenue'))}}
Add-Visual 'clusteredColumnChart' 'Revenue by Category' 15 435 590 255 @{Category=@{projections=@((New-ColumnProjection 'Category' 'SalesAnalytics.Category'))};Y=@{projections=@((New-MeasureProjection 'Total Revenue'))}}
Add-Visual 'clusteredBarChart' 'Revenue by Product' 625 435 605 255 @{Category=@{projections=@((New-ColumnProjection 'ProductName' 'SalesAnalytics.ProductName'))};Y=@{projections=@((New-MeasureProjection 'Total Revenue'))}}

$readme=@"
POWER BI ANALYTICS LAB
======================
SQL server: $SqlServer
Database: $Database
View: dbo.vw_SalesAnalytics
Verified SQL rows: $rows

1. Open SalesAnalytics.pbip with Power BI Desktop (64-bit).
2. Authenticate using a READ-ONLY Windows account and refresh the Import model.
3. Check the Sales Analytics page and all 7 visuals.
4. If any visuals are blank/unrecognized: the visual-role schema is version-sensitive.
   Save one working chart from Power BI Desktop as PBIP and use its visual.json
   as the reference. DO NOT overwrite your known-working project.
5. Save As .pbix after report renders.

SECURITY: uses Encrypt=True with TrustServerCertificate=True only in the
PowerShell SQL metadata check. Configure trusted SQL TLS certificates for
production. The M query uses Power BI's normal connector settings.

NOTICE: Script writes ONLY to this newly-created folder. Does not install
Power BI Desktop, alter PROD/DEV SQL databases, or modify SQL Agent jobs.
"@
Write-TextFile (Join-Path $run 'README.txt') $readme

$visualCount=@(Get-ChildItem -LiteralPath $page -Filter visual.json -Recurse -File).Count
Write-Host "`nCreated: $project" -ForegroundColor Green
Write-Host "PBIR visual JSON files: $visualCount" -ForegroundColor Green
Write-Host 'Power BI Desktop must open/refresh/validate the project before use.' -ForegroundColor Yellow
if($OpenWhenDone){Start-Process -FilePath $project}
