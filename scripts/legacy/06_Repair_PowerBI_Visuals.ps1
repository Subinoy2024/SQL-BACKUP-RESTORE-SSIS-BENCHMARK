#requires -Version 5.1
<#
Adds a fresh Sales Dashboard page to an existing PBIP report, preserving the old page.
Power BI Desktop must be completely closed first. Does not alter SQL or jobs.
Script-created PBIR report metadata must be validated in Power BI Desktop.
#>
[CmdletBinding()]
param([string]$ProjectFolder='C:\SQLLab\PowerBI_Dashboard')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest

$root=Join-Path $ProjectFolder 'AnalyticsLab.Report\definition'
$project=Join-Path $ProjectFolder 'AnalyticsLab.pbip'
$pagesRoot=Join-Path $root 'pages'
$pagesPath=Join-Path $pagesRoot 'pages.json'
if (-not (Test-Path -LiteralPath $project) -or -not (Test-Path -LiteralPath $pagesPath)) { throw "Project files not found in $ProjectFolder" }
if (Get-Process PBIDesktop -ErrorAction SilentlyContinue) { throw 'Close ALL Power BI Desktop windows before changing project files.' }

# Back up BOTH project artifacts, excluding prior backups.
$backupBase=Join-Path $ProjectFolder '_backups'
New-Item -ItemType Directory -Path $backupBase -Force | Out-Null
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$backup=Join-Path $backupBase "Before_Visual_Repair_$stamp"
New-Item -ItemType Directory -Path $backup -Force | Out-Null
Copy-Item -LiteralPath $project -Destination $backup
Copy-Item -LiteralPath (Join-Path $ProjectFolder 'AnalyticsLab.Report') -Destination $backup -Recurse
Write-Host "Backup saved to: $backup" -ForegroundColor Green

function Write-JsonFile([string]$path,$obj) {
    $parent=Split-Path -Parent $path
    if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $text=ConvertTo-Json -InputObject $obj -Depth 60
    [System.IO.File]::WriteAllText($path,$text,[System.Text.UTF8Encoding]::new($false))
}
function New-Id { 'visual' + [guid]::NewGuid().ToString('N').Substring(0,14) }
function Projection([string]$field,[bool]$measure) {
    $kind=if($measure){'Measure'}else{'Column'}
    $fieldNode=@{Expression=@{SourceRef=@{Entity='SalesAnalytics'}};Property=$field}
    return @{field=@{$kind=$fieldNode};queryRef="SalesAnalytics.$field";nativeQueryRef=$field}
}
function Add-Visual([string]$type,[int]$x,[int]$y,[int]$w,[int]$h,[string]$dim,[string]$measure,[int]$order) {
    $id=New-Id
    $state=[ordered]@{}
    if($type -eq 'card') {
        $state['Values']=@{projections=@((Projection $measure $true))}
    } elseif($type -eq 'tableEx') {
        $state['Values']=@{projections=@((Projection 'Region' $false),(Projection 'Category' $false),(Projection 'Total Revenue' $true))}
    } else {
        $state['Category']=@{projections=@((Projection $dim $false))}
        $state['Y']=@{projections=@((Projection $measure $true))}
    }
    $obj=[ordered]@{
        '$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.7.0/schema.json'
        name=$id
        position=@{x=$x;y=$y;z=(1000+$order);width=$w;height=$h;tabOrder=$order}
        visual=@{visualType=$type;query=@{queryState=$state};drillFilterOtherVisuals=$true}
    }
    Write-JsonFile (Join-Path $visualRoot "$id\visual.json") $obj
}

$pages=Get-Content -LiteralPath $pagesPath -Raw | ConvertFrom-Json
$pageName='SalesDashboard'+([guid]::NewGuid().ToString('N').Substring(0,10))
$pageDir=Join-Path $pagesRoot $pageName
$visualRoot=Join-Path $pageDir 'visuals'
New-Item -ItemType Directory -Path $visualRoot -Force | Out-Null
Write-JsonFile (Join-Path $pageDir 'page.json') ([ordered]@{
    '$schema'='https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json'
    name=$pageName;displayName='Sales Dashboard';displayOption='FitToPage';height=720;width=1280
})
Add-Visual 'card' 20 20 390 125 '' 'Total Revenue' 1
Add-Visual 'card' 445 20 390 125 '' 'Total Orders' 2
Add-Visual 'card' 870 20 390 125 '' 'Customers Served' 3
Add-Visual 'lineChart' 20 175 790 250 'OrderDate' 'Total Revenue' 4
Add-Visual 'clusteredBarChart' 840 175 420 250 'Region' 'Total Revenue' 5
Add-Visual 'clusteredColumnChart' 20 455 600 240 'Category' 'Total Revenue' 6
Add-Visual 'clusteredBarChart' 645 455 615 240 'ProductName' 'Total Revenue' 7
$existing=@($pages.pageOrder)
$pages.pageOrder=@($existing + @($pageName))
$pages.activePageName=$pageName
Write-JsonFile $pagesPath $pages
Write-Host "New page: Sales Dashboard ($pageName)" -ForegroundColor Green
Write-Host "Seven visuals created. Open: $project" -ForegroundColor Green
Write-Warning 'Validate visuals in Power BI Desktop. If the new page is still blank, PBIR compatibility needs Desktop-exported reference visuals. Revert by restoring the backup folder.'
