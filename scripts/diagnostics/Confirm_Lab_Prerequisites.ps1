#requires -Version 5.1
# Read-only preflight for the SQL/Power BI/SSIS lab.
param([string]$Root='C:\SQLLab\SSIS')
$files=@(
 @{Label='Excel workbook';Path=(Join-Path $Root 'Input\SalesBenchmark_100K.xlsx')},
 @{Label='SSIS package';Path=(Join-Path $Root 'SSIS_Performance_Lab\SalesExcel_100K.dtsx')},
 @{Label='DTExec x64';Path='C:\Program Files\Microsoft SQL Server\170\DTS\Binn\DTExec.exe'},
 @{Label='DTExec x86';Path='C:\Program Files (x86)\Microsoft SQL Server\170\DTS\Binn\DTExec.exe'},
 @{Label='Visual Studio';Path='C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\devenv.exe'}
)
$files | ForEach-Object {[pscustomobject]@{Item=$_.Label;Exists=(Test-Path -LiteralPath $_.Path);Path=$_.Path}} | Format-Table -Wrap
$reg='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if(Test-Path $reg){Get-ItemProperty $reg | Select-Object Platform,ProductReleaseIds,VersionToReport | Format-List}
