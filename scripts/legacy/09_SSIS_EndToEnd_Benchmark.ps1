#requires -Version 5.1
<#
.SYNOPSIS
    End-to-end SSIS Excel benchmark orchestrator for the isolated DEV lab.
.DESCRIPTION
    Validates runtime, ACE provider, input workbook, package, database/table,
    executes a real DTExec package in selected architecture(s), checks row counts,
    and exports run CSV and summary CSV. It does NOT create a DTSX package.
    No PROD database changes. Does not create or alter scheduled jobs.
.EXAMPLE
    .\09_SSIS_EndToEnd_Benchmark.ps1 -Discover
.EXAMPLE
    .\09_SSIS_EndToEnd_Benchmark.ps1 -PackagePath 'C:\SQLLab\SSIS\SSIS_Performance_Lab\SalesExcel_100K.dtsx' -Architecture x64 -Runs 3 -AllowResetBenchmarkTable
#>
[CmdletBinding()]
param(
 [string]$PackagePath = '',
 [ValidateSet('x64','x86','Both')][string]$Architecture = 'x64',
 [ValidateRange(1,20)][int]$Runs = 3,
 [string]$ExcelFile = 'C:\SQLLab\SSIS\Input\SalesBenchmark_100K.xlsx',
 [string]$SqlServer = 'localhost',
 [string]$Database = 'PowerBIAnalytics_DEV',
 [string]$OutputDirectory = 'C:\SQLLab\SSIS\Results',
 [int]$ExpectedRows = 100000,
 [switch]$AllowResetBenchmarkTable,
 [switch]$Discover
)
$ErrorActionPreference = 'Stop'
$dt = @{
 x64 = 'C:\Program Files\Microsoft SQL Server\170\DTS\Binn\DTExec.exe'
 x86 = 'C:\Program Files (x86)\Microsoft SQL Server\170\DTS\Binn\DTExec.exe'
}
$ps = @{
 x64 = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
 x86 = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
}
function Assert-File([string]$p) { if (!(Test-Path -LiteralPath $p -PathType Leaf)) { throw "Required file not found: $p" } }
function Get-DbScalar([string]$sql) {
 $cs = "Server=$SqlServer;Database=$Database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15"
 $c = [System.Data.SqlClient.SqlConnection]::new($cs)
 try { $c.Open(); $cmd=$c.CreateCommand();$cmd.CommandText=$sql;$cmd.CommandTimeout=120;return $cmd.ExecuteScalar() }
 finally { $c.Dispose() }
}
function Test-Ace([string]$a) {
 $probe = @'
$p='Microsoft.ACE.OLEDB.16.0'
$c=[System.Data.OleDb.OleDbConnection]::new("Provider=$p;Data Source=C:\nonexistent_ACE_probe.xlsx;Extended Properties=`"Excel 12.0 Xml;HDR=YES`";")
try { $c.Open(); exit 0 } catch { if ($_.Exception.ToString() -match 'not registered|not installed|cannot be found') { exit 21 }; exit 0 } finally { $c.Dispose() }
'@
 $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probe))
 & $ps[$a] -NoProfile -NonInteractive -EncodedCommand $encoded 2>$null | Out-Null
 return ($LASTEXITCODE -eq 0)
}
Write-Host "`n=== DEV SSIS BENCHMARK PREFLIGHT ===" -ForegroundColor Cyan
if ($env:COMPUTERNAME -ne 'DCSQLDB-SEC1') { Write-Warning "Host is $env:COMPUTERNAME (expected DCSQLDB-SEC1). Ensure this is NOT PROD." }
if ($Database -ne 'PowerBIAnalytics_DEV') { throw 'Safety stop: only PowerBIAnalytics_DEV is allowed.' }
Assert-File $ExcelFile
foreach($a in @('x86','x64')) {
 $runtime = Test-Path -LiteralPath $dt[$a] -PathType Leaf
 $ace = if ($runtime) { Test-Ace $a } else { $false }
 Write-Host "$a : DTExec=$runtime ACE16Load=$ace"
}
$tablePresent = [int](Get-DbScalar "SELECT CASE WHEN OBJECT_ID('dbo.SSIS_Benchmark_Results','U') IS NULL THEN 0 ELSE 1 END")
if ($tablePresent -eq 0) { throw 'Benchmark table missing; DEV daily refresh may have removed it. Recreate lab table before running.' }
$cols = [int](Get-DbScalar "SELECT COUNT(*) FROM sys.columns WHERE object_id=OBJECT_ID('dbo.SSIS_Benchmark_Results') AND name IN ('OrderID','ProductID','Quantity','SalesAmount')")
if ($cols -ne 4) { throw 'Benchmark table does not have all four required columns.' }
if ($Discover) {
 Write-Host 'Package discovery:'
 Get-ChildItem 'C:\SQLLab\SSIS' -Filter '*.dtsx' -Recurse -File -ErrorAction SilentlyContinue | Select-Object FullName,Length,LastWriteTime | Format-Table -AutoSize
 Write-Host 'Preflight complete; no package executed.'
 return
}
if (-not $PackagePath) { throw 'Pass -PackagePath pointing to a genuine SSIS .dtsx created with Excel Source -> OLE DB Destination. Run -Discover first.' }
Assert-File $PackagePath
if (-not $AllowResetBenchmarkTable) { throw 'Safety stop: reruns require resetting the lab destination. Re-run with -AllowResetBenchmarkTable after confirming it has no data to retain.' }
$arches=if ($Architecture -eq 'Both') {@('x86','x64')} else {@($Architecture)}
foreach($a in $arches) {
 Assert-File $dt[$a]
 if (!(Test-Ace $a)) { throw "$a ACE provider is not operational. Install a compatible $a runtime or use an isolated matching test VM. Benchmark not run." }
}
# Avoid unattended DEV refresh colliding with benchmark.
$now=Get-Date
if ($now.TimeOfDay -ge [TimeSpan]::FromHours(4.5) -and $now.TimeOfDay -le [TimeSpan]::FromHours(6.5)) {
 throw 'Safety stop: current time is within 04:30-06:30 SQL backup/DEV refresh window.'
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$rows=[System.Collections.Generic.List[object]]::new()
foreach($a in $arches) {
 for($i=1;$i -le $Runs;$i++) {
  Write-Host "`n=== $a RUN $i/$Runs ===" -ForegroundColor Cyan
  # Only reset the dedicated lab table, never existing analytics business tables.
  $cs="Server=$SqlServer;Database=$Database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True"
  $con=[System.Data.SqlClient.SqlConnection]::new($cs)
  try { $con.Open();$cmd=$con.CreateCommand();$cmd.CommandText='TRUNCATE TABLE dbo.SSIS_Benchmark_Results';$cmd.ExecuteNonQuery() | Out-Null }
  finally {$con.Dispose()}
  $log=Join-Path $OutputDirectory "SSIS_${stamp}_${a}_run${i}.log"
  $started=Get-Date
  # A file-system .dtsx package: Windows service identity must be allowed to write to SQL target.
  $output = & $dt[$a] /F $PackagePath /REPORTING E 2>&1
  $exit=$LASTEXITCODE
  $elapsed=((Get-Date)-$started).TotalSeconds
  $output | Out-File -LiteralPath $log -Encoding UTF8
  $count=[int](Get-DbScalar 'SELECT COUNT_BIG(*) FROM dbo.SSIS_Benchmark_Results')
  $passed=($exit -eq 0 -and $count -eq $ExpectedRows)
  $rows.Add([pscustomobject]@{ Architecture=$a; Run=$i; Seconds=[math]::Round($elapsed,3); ExitCode=$exit; Rows=$count; Expected=$ExpectedRows; RowsPerSecond=if($elapsed -gt 0){[math]::Round($count/$elapsed,2)}else{0}; Passed=$passed; Log=$log })
  Write-Host "Exit=$exit Rows=$count Seconds=$([math]::Round($elapsed,2)) Passed=$passed"
  if (-not $passed) { break }
 }
}
$runCsv=Join-Path $OutputDirectory "SSIS_Runs_$stamp.csv"
$rows | Export-Csv -LiteralPath $runCsv -NoTypeInformation -Encoding UTF8
$summary = @($rows | Group-Object Architecture | ForEach-Object {
 $good=@($_.Group | Where-Object Passed | Sort-Object Seconds)
 [pscustomobject]@{Architecture=$_.Name; SuccessfulRuns=$good.Count; MedianSeconds=if($good.Count){$good[[int][math]::Floor(($good.Count-1)/2)].Seconds}else{$null}; AllPassed=(@($_.Group | Where-Object { -not $_.Passed }).Count -eq 0 -and $_.Group.Count -eq $Runs)}
})
$summaryCsv=Join-Path $OutputDirectory "SSIS_Summary_$stamp.csv"
$summary | Export-Csv -LiteralPath $summaryCsv -NoTypeInformation -Encoding UTF8
$summary | Format-Table -AutoSize
Write-Host "`nRun results: $runCsv`nSummary: $summaryCsv"
if (@($rows | Where-Object { -not $_.Passed }).Count) { throw 'One or more SSIS runs failed or loaded the wrong number of rows. Review log files.' }
