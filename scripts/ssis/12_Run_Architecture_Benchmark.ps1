#requires -Version 5.1
<#
Runs an EXISTING, validated .dtsx package with the chosen DTExec architecture.
WARNING: Truncates only dbo.SSIS_Benchmark_Results, with explicit opt-in.
The package must target the specified lab table; inspect it before execution.
Avoid the 04:30-06:30 scheduled DEV restore window.
#>
[CmdletBinding()]
param(
 [ValidateSet('x86','x64')][string]$Architecture='x86',
 [string]$PackagePath='C:\SQLLab\SSIS\SSIS_Performance_Lab\SalesExcel_100K.dtsx',
 [string]$Server='localhost',
 [string]$Database='PowerBIAnalytics_DEV',
 [string]$OutputDirectory='C:\SQLLab\SSIS\Results',
 [ValidateRange(1,20)][int]$Runs=3,
 [ValidateRange(1,100000000)][int]$ExpectedRows=100000,
 [switch]$AllowResetBenchmarkTable
)
$ErrorActionPreference='Stop'
if ($Database -ne 'PowerBIAnalytics_DEV') { throw 'Safety stop: this runner only supports PowerBIAnalytics_DEV. Review and adapt the source for other labs.' }
if (-not $AllowResetBenchmarkTable) { throw 'This test truncates dbo.SSIS_Benchmark_Results. Pass -AllowResetBenchmarkTable to consent.' }
$dtExec=if($Architecture -eq 'x86'){'C:\Program Files (x86)\Microsoft SQL Server\170\DTS\Binn\DTExec.exe'}else{'C:\Program Files\Microsoft SQL Server\170\DTS\Binn\DTExec.exe'}
foreach($p in @($PackagePath,$dtExec)){if(-not(Test-Path -LiteralPath $p -PathType Leaf)){throw "Missing file: $p"}}
if ((Get-Date).TimeOfDay -ge [TimeSpan]::FromHours(4.5) -and (Get-Date).TimeOfDay -le [TimeSpan]::FromHours(6.5)) {throw 'Scheduled DEV restore window 04:30-06:30: run outside this period.'}
$cs="Server=$Server;Database=$Database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15"
function Invoke-Scalar([string]$sql){
 $c=[System.Data.SqlClient.SqlConnection]::new($cs)
 try{$c.Open();$cmd=$c.CreateCommand();$cmd.CommandText=$sql;$cmd.CommandTimeout=180;return $cmd.ExecuteScalar()}finally{$c.Dispose()}
}
function Invoke-NonQuery([string]$sql){
 $c=[System.Data.SqlClient.SqlConnection]::new($cs)
 try{$c.Open();$cmd=$c.CreateCommand();$cmd.CommandText=$sql;$cmd.CommandTimeout=180;[void]$cmd.ExecuteNonQuery()}finally{$c.Dispose()}
}
if (-not (Invoke-Scalar "SELECT CASE WHEN OBJECT_ID('dbo.SSIS_Benchmark_Results','U') IS NULL THEN 0 ELSE 1 END")){throw 'Lab benchmark table missing; restore job may have replaced DEV database.'}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
Write-Host "=== VALIDATE $Architecture PACKAGE ===" -ForegroundColor Cyan
& $dtExec /F $PackagePath /Validate /REPORTING E
if($LASTEXITCODE -ne 0){throw "DTExec validation failed: $LASTEXITCODE"}
$results=@()
for($i=1;$i -le $Runs;$i++){
 Write-Host "=== $Architecture RUN $i/$Runs ===" -ForegroundColor Cyan
 Invoke-NonQuery 'TRUNCATE TABLE dbo.SSIS_Benchmark_Results'
 $log=Join-Path $OutputDirectory "SSIS_${Architecture}_${stamp}_run${i}.log"
 $timer=[Diagnostics.Stopwatch]::StartNew()
 & $dtExec /F $PackagePath /REPORTING E *>&1 | Out-File -LiteralPath $log -Encoding UTF8
 $exit=$LASTEXITCODE
 $timer.Stop()
 $count=[long](Invoke-Scalar 'SELECT COUNT_BIG(*) FROM dbo.SSIS_Benchmark_Results')
 $pass=($exit -eq 0 -and $count -eq $ExpectedRows)
 $item=[pscustomobject]@{Architecture=$Architecture;Run=$i;Seconds=[math]::Round($timer.Elapsed.TotalSeconds,3);Rows=$count;ExitCode=$exit;Passed=$pass;Log=$log}
 $results+= $item
 Write-Host "Run $i : $($item.Seconds) sec | Rows=$count | Exit=$exit | Passed=$pass"
 if(-not $pass){break}
}
$csv=Join-Path $OutputDirectory "SSIS_${Architecture}_${stamp}.csv"
$results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
Write-Host "CSV: $csv"
if (@($results | Where-Object {-not $_.Passed}).Count) {throw 'Benchmark failed: see log. CSV contains partial results.'}
$ordered=@($results | Sort-Object Seconds | Select-Object -ExpandProperty Seconds)
$median=if($ordered.Count%2){$ordered[[int][math]::Floor($ordered.Count/2)]}else{($ordered[$ordered.Count/2-1]+$ordered[$ordered.Count/2])/2}
$avg=($results | Measure-Object -Property Seconds -Average).Average
Write-Host ("$Architecture median={0:N3} s, average={1:N3} s; successful runs={2}" -f $median,$avg,$results.Count) -ForegroundColor Green
