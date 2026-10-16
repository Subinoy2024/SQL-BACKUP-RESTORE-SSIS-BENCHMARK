#requires -Version 5.1
<#
Creates a Power BI Desktop project (.pbip) with a SQL-backed Import semantic model
and preconfigured sales dashboard visuals. Run on a Windows machine with Power BI Desktop.
Does NOT modify SQL Server. PBIP/PBIR/TMDL compatibility should be validated in Desktop.
#>
[CmdletBinding()]
param(
    [string]$Server = '192.168.88.18',
    [string]$Database = 'PowerBIAnalytics_PROD',
    [string]$OutputFolder = 'C:\SQLLab\PowerBI_Dashboard',
    [switch]$SkipSqlCheck
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ($Server -notmatch '^[a-zA-Z0-9_.\\,-]+$') { throw 'Invalid SQL Server name.' }
if ($Database -notmatch '^[a-zA-Z0-9_]+$') { throw 'Invalid database name.' }
if (-not $SkipSqlCheck) {
    $connection = [System.Data.SqlClient.SqlConnection]::new("Server=$Server;Database=$Database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=10")
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = "SELECT COUNT_BIG(*) FROM dbo.vw_SalesAnalytics"
        $rows = [long]$command.ExecuteScalar()
        Write-Host "Confirmed $rows reporting rows in $Server / $Database" -ForegroundColor Green
    } finally { $connection.Dispose() }
}
$projectName='AnalyticsLab'
$proj = Join-Path $OutputFolder ($projectName+'.pbip')
$report = Join-Path $OutputFolder ($projectName+'.Report')
$model = Join-Path $OutputFolder ($projectName+'.SemanticModel')
if (Test-Path $proj) { throw "Project already exists at $proj. Select another OutputFolder to avoid overwriting work." }
function Write-Utf8($Path,[string]$Value) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path,$Value,$utf8)
}
function Write-Json($Path,$Object) { Write-Utf8 $Path ($Object | ConvertTo-Json -Depth 45) }
function New-Id { ([guid]::NewGuid().ToString('N').Substring(0,20)) }
$schemaRoot = 'https://developer.microsoft.com/json-schemas/fabric/'
Write-Json $proj ([ordered]@{
    '$schema'=$schemaRoot+'pbip/pbipProperties/1.0.0/schema.json';
    version='1.0';artifacts=@(@{report=@{path=($projectName+'.Report')}})
})
Write-Json (Join-Path $report 'definition.pbir') ([ordered]@{
    '$schema'=$schemaRoot+'item/report/definitionProperties/2.0.0/schema.json';version='4.0';
    datasetReference=@{byPath=@{path='../AnalyticsLab.SemanticModel'}}
})
Write-Json (Join-Path $model 'definition.pbism') ([ordered]@{
    version='4.0';settings=@{}
})
# TMDL files: tabs are significant. This model defines one import table and six measures.
Write-Utf8 (Join-Path $model 'definition/database.tmdl') "database AnalyticsLab`n`tcompatibilityLevel: 1601`n"
Write-Utf8 (Join-Path $model 'definition/model.tmdl') "model Model`n`tculture: en-US`n`tdefaultPowerBIDataSourceVersion: powerBI_V3`n`tdataAccessOptions`n`t`tlegacyRedirects`n`t`treturnErrorValuesAsNull`n"
$tab=[char]9
$columns=@(
    @('OrderID','int64'),@('OrderDate','dateTime'),@('OrderStatus','string'),
    @('CustomerID','int64'),@('CustomerName','string'),@('Region','string'),
    @('ProductID','int64'),@('ProductName','string'),@('Category','string'),
    @('Quantity','int64'),@('UnitPrice','decimal'),@('SalesAmount','decimal')
)
$lines=[System.Collections.Generic.List[string]]::new()
$lines.Add('table SalesAnalytics')
$lines.Add('')
$measures=@(
    @('Total Revenue','SUM(SalesAnalytics[SalesAmount])','#,##0.00'),
    @('Total Orders','DISTINCTCOUNT(SalesAnalytics[OrderID])','#,##0'),
    @('Total Quantity','SUM(SalesAnalytics[Quantity])','#,##0'),
    @('Average Order Value','DIVIDE([Total Revenue], [Total Orders])','#,##0.00'),
    @('Customers Served','DISTINCTCOUNT(SalesAnalytics[CustomerID])','#,##0'),
    @('Average Unit Price','AVERAGE(SalesAnalytics[UnitPrice])','#,##0.00')
)
foreach($m in $measures) {
    $lines.Add("${tab}measure '$($m[0])' = $($m[1])")
    $lines.Add("${tab}${tab}formatString: $($m[2])")
    $lines.Add('')
}
foreach ($c in $columns) {
    $lines.Add("${tab}column $($c[0])")
    $lines.Add("${tab}${tab}dataType: $($c[1])")
    $lines.Add("${tab}${tab}summarizeBy: none")
    $lines.Add("${tab}${tab}sourceColumn: $($c[0])")
    $lines.Add('')
}
$lines.Add("${tab}partition SalesAnalytics = m")
$lines.Add("${tab}${tab}mode: import")
$lines.Add("${tab}${tab}source =")
$lines.Add("${tab}${tab}${tab}let")
$lines.Add("${tab}${tab}${tab}${tab}Source = Sql.Database(`"$Server`", `"$Database`"),")
$lines.Add("${tab}${tab}${tab}${tab}Analytics = Source{[Schema=`"dbo`",Item=`"vw_SalesAnalytics`"]}[Data]")
$lines.Add("${tab}${tab}${tab}in")
$lines.Add("${tab}${tab}${tab}${tab}Analytics")
Write-Utf8 (Join-Path $model 'definition/tables/SalesAnalytics.tmdl') ($lines -join "`r`n")
Write-Json (Join-Path $report 'definition/version.json') ([ordered]@{
    '$schema'=$schemaRoot+'item/report/definition/versionMetadata/1.0.0/schema.json';version='2.0.0'
})
# Minimal report definition with system theme omitted. Desktop may enrich defaults on first save.
Write-Json (Join-Path $report 'definition/report.json') ([ordered]@{
    '$schema'=$schemaRoot+'item/report/definition/report/3.2.0/schema.json'
})
$page = New-Id
Write-Json (Join-Path $report 'definition/pages/pages.json') ([ordered]@{
    '$schema'=$schemaRoot+'item/report/definition/pagesMetadata/1.0.0/schema.json'
    pageOrder=@($page);activePageName=$page
})
Write-Json (Join-Path $report "definition/pages/$page/page.json") ([ordered]@{
    '$schema'=$schemaRoot+'item/report/definition/page/2.1.0/schema.json';name=$page;
    displayName='Sales Analytics';displayOption='FitToPage';height=720;width=1280
})
function Get-Field([string]$Name,[bool]$IsMeasure) {
    $kind = if($IsMeasure){'Measure'}else{'Column'}
    return @{ $kind = @{Expression=@{SourceRef=@{Entity='SalesAnalytics'}};Property=$Name} }
}
function Get-Proj([string]$Name,[bool]$IsMeasure) {
    return @{field=(Get-Field $Name $IsMeasure);queryRef="SalesAnalytics.$Name"}
}
function Add-Visual([string]$Type,[int]$X,[int]$Y,[int]$W,[int]$H,[string]$Dimension,[string]$Measure) {
    $name=New-Id
    $q=[ordered]@{}
    if ($Type -eq 'cardVisual') { $q['Data'] = @{projections=@((Get-Proj $Measure $true))} }
    else {
        $q['Category'] = @{projections=@((Get-Proj $Dimension $false))}
        $q['Y'] = @{projections=@((Get-Proj $Measure $true))}
    }
    Write-Json (Join-Path $report "definition/pages/$page/visuals/$name/visual.json") ([ordered]@{
        '$schema'=$schemaRoot+'item/report/definition/visualContainer/2.7.0/schema.json';
        name=$name;position=@{x=$X;y=$Y;z=0;width=$W;height=$H;tabOrder=0};
        visual=@{visualType=$Type;query=@{queryState=$q}}
    })
}
Add-Visual 'cardVisual' 20 20 390 130 '' 'Total Revenue'
Add-Visual 'cardVisual' 440 20 390 130 '' 'Total Orders'
Add-Visual 'cardVisual' 860 20 390 130 '' 'Customers Served'
Add-Visual 'lineChart' 20 170 790 275 'OrderDate' 'Total Revenue'
Add-Visual 'barChart' 830 170 420 275 'Region' 'Total Revenue'
Add-Visual 'columnChart' 20 465 600 235 'Category' 'Total Revenue'
Add-Visual 'barChart' 640 465 610 235 'ProductName' 'Total Revenue'
$readme=@"
Power BI Desktop project - AnalyticsLab
=======================================
Server: $Server
Database: $Database
View: dbo.vw_SalesAnalytics

1. Install/update Power BI Desktop (Windows).
2. Open $proj using Power BI Desktop.
3. When prompted, use Windows credentials for a SQL read-only user.
4. Select Home > Refresh to load SQL data.
5. Review visuals, then File > Save As > Power BI report (.pbix) if desired.

The PBIP/PBIR files are generated rather than exported by Power BI Desktop;
Power BI may ask to upgrade/repair their metadata or a visual type.
The project cannot be runtime-tested on the generator machine unless Power BI Desktop is installed.

DAX: Total Revenue, Total Orders, Total Quantity, Average Order Value,
Customers Served, Average Unit Price.

This script only reads SQL (unless -SkipSqlCheck) and writes files to OutputFolder.
No cloud publication, Pro licensing, or credentials are embedded.
"@
Write-Utf8 (Join-Path $OutputFolder 'README_PowerBI.txt') $readme
Write-Host "Project created: $proj" -ForegroundColor Green
Write-Host 'Open in Power BI Desktop, authenticate, Refresh, and Save As .pbix.'
