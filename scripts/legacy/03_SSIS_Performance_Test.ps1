#requires -Version 5.1
<#
 Run on DEV/ETL host with licensed/supported SSIS runtime installed and actual .dtsx packages.
 Does NOT manufacture results: if either runtime is missing the script fails with an explanation.
 Example: .\03_SSIS_Performance_Test.ps1 -PackageDirectory 'C:\SQLLab\SSIS' -Runs 3
 Use packages with identical data and connection settings for fair 32/64-bit comparison.
 #>
[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$PackageDirectory,
 [string]$OutputDirectory='C:\SQLLab\Benchmarks',
 [int]$Runs=3,
 [int]$TimeoutMinutes=180,
 [string]$Dtexec64='',
 [string]$Dtexec32=''
)
$ErrorActionPreference='Stop'
if($Runs -lt 1 -or $Runs -gt 20){throw 'Runs must be 1..20.'}
if(-not(Test-Path -LiteralPath $PackageDirectory -PathType Container)){throw "Package directory missing: $PackageDirectory"}
$packages=@(Get-ChildItem -LiteralPath $PackageDirectory -File -Filter '*.dtsx')
if($packages.Count -eq 0){throw 'No .dtsx packages found. Supply real Excel/SSIS packages (not fabricated workloads).'}
function Find-Dtexec([bool]$bit64){
 $root=if($bit64){${env:ProgramFiles}}else{${env:ProgramFiles(x86)}}
 if(-not $root){return $null}
 foreach($ver in @('170','160','150','140','130','120','110')){
  $p=Join-Path $root "Microsoft SQL Server\$ver\DTS\Binn\DTExec.exe"
  if(Test-Path -LiteralPath $p){return $p}
 }
 return $null
}
if(-not $Dtexec64){$Dtexec64=Find-Dtexec $true}
if(-not $Dtexec32){$Dtexec32=Find-Dtexec $false}
Write-Host "64-bit DTExec: $Dtexec64";Write-Host "32-bit DTExec: $Dtexec32"
if(-not $Dtexec64 -or -not (Test-Path -LiteralPath $Dtexec64)){throw '64-bit DTExec not found. Install compatible Integration Services runtime.'}
if(-not $Dtexec32 -or -not (Test-Path -LiteralPath $Dtexec32)){throw '32-bit DTExec not found. Install supported 32-bit SSIS components or specify -Dtexec32.'}
Write-Host "`nACE registry probes (these are hints, not proof a package provider can load):" -ForegroundColor Cyan
foreach($view in @('Registry64','Registry32')){
 $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::$view)
 try{$key=$base.OpenSubKey('SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0')
  Write-Host "${view}: $(if($key){'ACE 12 ProgID registered'}else{'ACE 12 ProgID not found'})"
 } finally {$base.Dispose()}
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$results=[System.Collections.Generic.List[object]]::new()
$runStamp=Get-Date -Format 'yyyyMMdd_HHmmss'
foreach($pkg in $packages){
 foreach($round in 1..$Runs){
  foreach($architecture in @('32-bit','64-bit')){
   $exe=if($architecture -eq '32-bit'){$Dtexec32}else{$Dtexec64}
   $label=('{0}_{1}_{2}_{3}' -f $pkg.BaseName,$architecture,$round,$runStamp) -replace '[^a-zA-Z0-9_\-]','_'
   $stdout=Join-Path $OutputDirectory "$label.stdout.log"
   $stderr=Join-Path $OutputDirectory "$label.stderr.log"
   Write-Host "Running $($pkg.Name), $architecture, repetition $round of $Runs" -ForegroundColor Cyan
   $sw=[System.Diagnostics.Stopwatch]::StartNew();$exitCode=$null;$status='FAILED'
   try{
     $args='/F "'+$pkg.FullName+'" /REPORTING E'
     $p=Start-Process -FilePath $exe -ArgumentList $args -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr -WindowStyle Hidden
     if(-not $p.WaitForExit($TimeoutMinutes*60000)){
       try{$p.Kill()}catch{}
       $status='TIMEOUT';$exitCode=-1
     }else{
       $exitCode=$p.ExitCode
       $status=if($exitCode -eq 0){'SUCCESS'}else{'FAILED'}
     }
   }catch{
     $_.Exception.Message | Add-Content -Path $stderr
     $exitCode=-2
   }finally{$sw.Stop()}
   $results.Add([pscustomobject]@{Package=$pkg.Name;Architecture=$architecture;Run=$round;Status=$status;ExitCode=$exitCode;DurationSeconds=[math]::Round($sw.Elapsed.TotalSeconds,3);Stdout=$stdout;Stderr=$stderr})
  }
 }
}
$csv=Join-Path $OutputDirectory "SSIS_RunDetails_$runStamp.csv"
$results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
$summary=@(foreach($pkg in $packages){
 $ok32=@($results | Where-Object {$_.Package -eq $pkg.Name -and $_.Architecture -eq '32-bit' -and $_.Status -eq 'SUCCESS'})
 $ok64=@($results | Where-Object {$_.Package -eq $pkg.Name -and $_.Architecture -eq '64-bit' -and $_.Status -eq 'SUCCESS'})
 $avg32=if($ok32.Count -eq $Runs){($ok32 | Measure-Object DurationSeconds -Average).Average}else{$null}
 $avg64=if($ok64.Count -eq $Runs){($ok64 | Measure-Object DurationSeconds -Average).Average}else{$null}
 $gain=if($null -ne $avg32 -and $null -ne $avg64 -and $avg32 -gt 0){[math]::Round((($avg32-$avg64)/$avg32)*100,2)}else{$null}
 [pscustomobject]@{Package=$pkg.Name;Runs32Successful=$ok32.Count;Runs64Successful=$ok64.Count;Avg32Seconds=$avg32;Avg64Seconds=$avg64;PerformanceImprovementPct=$gain}
})
$summaryCsv=Join-Path $OutputDirectory "SSIS_Comparison_$runStamp.csv"
$summary | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
$summary | Format-Table -AutoSize
Write-Host "Run details: $csv`nComparison: $summaryCsv" -ForegroundColor Green
Write-Warning 'Run tests against isolated DEV/synthetic data with consistent dataset, Excel drivers, package settings, cache conditions and no competing jobs. Validate outputs/row counts separately.'
