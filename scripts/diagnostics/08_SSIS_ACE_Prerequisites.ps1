#requires -Version 5.1
# SSIS/ACE preflight: read-only, run on DEV 192.168.88.19
$ErrorActionPreference = 'Continue'
Write-Host "=== HOST ===" -ForegroundColor Cyan
Get-CimInstance Win32_OperatingSystem | Select-Object Caption,OSArchitecture,Version | Format-List
Write-Host "=== SQL / SSIS SERVICES ===" -ForegroundColor Cyan
Get-CimInstance Win32_Service | Where-Object { $_.Name -match 'MSSQL|SQLSERVERAGENT|MsDtsServer|SSIS' } | Select-Object Name,State,StartName | Format-Table -AutoSize
Write-Host "=== SSIS EXECUTABLES (x86/x64) ===" -ForegroundColor Cyan
$roots = @("$env:ProgramFiles\Microsoft SQL Server", "${env:ProgramFiles(x86)}\Microsoft SQL Server")
$tools = foreach($root in $roots) {
 if(Test-Path $root) {
  Get-ChildItem -Path $root -Filter DTExec.exe -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
   [pscustomobject]@{Path=$_.FullName; Architecture=if($_.FullName -like '*Program Files (x86)*'){'x86'}else{'x64'}; Version=$_.VersionInfo.FileVersion}
  }
 }
}
if($tools) {$tools | Sort-Object Path | Format-Table -AutoSize} else {Write-Warning 'No DTExec.exe found. SSIS runtime may not be installed.'}
Write-Host "=== ACE OLE DB REGISTRATION (both registry views) ===" -ForegroundColor Cyan
foreach($arch in @('Registry64','Registry32')) {
 $view = [Microsoft.Win32.RegistryView]::$arch
 $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot,$view)
 try {
  foreach($provider in @('Microsoft.ACE.OLEDB.12.0','Microsoft.ACE.OLEDB.16.0')) {
   $key = $base.OpenSubKey("$provider\CLSID")
   [pscustomobject]@{Architecture=$arch; Provider=$provider; Registered=[bool]$key; CLSID=if($key){$key.GetValue('')}else{''}}
   if($key){$key.Dispose()}
  }
 } finally {$base.Dispose()}
}
Write-Host "=== INSTALLED ACE/SSIS SOFTWARE (registry uninstall keys) ===" -ForegroundColor Cyan
$uninstall = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
Get-ItemProperty -Path $uninstall -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -match 'Access Database Engine|Access Runtime|Integration Services|SQL Server Data Tools'} | Select-Object DisplayName,DisplayVersion,Publisher | Sort-Object DisplayName -Unique | Format-Table -AutoSize
Write-Host "=== PACKAGE INVENTORY (no execution) ===" -ForegroundColor Cyan
$packageDirs = @('C:\SQLLab\SSIS','C:\SSIS','C:\ETL')
foreach($dir in $packageDirs){if(Test-Path $dir){Get-ChildItem $dir -Filter *.dtsx -Recurse -File -ErrorAction SilentlyContinue | Select-Object FullName,Length,LastWriteTime}}
Write-Host "=== DEV SQL CHECK ===" -ForegroundColor Cyan
try {
 $cs='Server=localhost;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=10'
 $cn=New-Object System.Data.SqlClient.SqlConnection($cs)
 $cn.Open()
 $cmd=$cn.CreateCommand()
 $cmd.CommandText="SELECT @@SERVERNAME AS ServerName, SERVERPROPERTY('Edition') AS Edition, DB_ID('PowerBIAnalytics_DEV') AS DevDatabaseID"
 $r=$cmd.ExecuteReader(); $tbl=New-Object System.Data.DataTable; $tbl.Load($r); $tbl | Format-Table -AutoSize
 $cn.Close()
} catch {Write-Warning "SQL inspection failed: $($_.Exception.Message)"} finally {if($cn){$cn.Dispose()}}
Write-Host 'PRECHECK COMPLETE (no changes made)' -ForegroundColor Green
