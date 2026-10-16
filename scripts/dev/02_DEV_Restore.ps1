#requires -Version 5.1
<#
 Run on 192.168.88.19 as DCCLOUD\vmadmin (or an authorized administrator).
 Save THIS file to C:\SQLLab\02_DEV_Restore.ps1 before -InstallJob.
 Example: .\02_DEV_Restore.ps1 -RunOnce -ApproveOverwrite
          .\02_DEV_Restore.ps1 -InstallJob -ApproveOverwrite
 The restore is allowed ONLY for an existing database marked SYNTHETIC-POWERBI-LAB.
 #>
[CmdletBinding()]
param(
 [string]$Server='localhost',
 [string]$Share='\\192.168.88.18\DBBACKUP$',
 [string]$SourceDatabase='PowerBIAnalytics_PROD',
 [string]$TargetDatabase='PowerBIAnalytics_DEV',
 [switch]$RunOnce,
 [switch]$InstallJob,
 [switch]$ApproveOverwrite
)
$ErrorActionPreference='Stop'
if($SourceDatabase -ne 'PowerBIAnalytics_PROD' -or $TargetDatabase -ne 'PowerBIAnalytics_DEV'){throw 'Restricted to the fixed lab database names.'}
if(-not ($RunOnce -or $InstallJob)){throw 'Specify -RunOnce or -InstallJob.'}
if(-not $ApproveOverwrite){throw 'Explicit -ApproveOverwrite required; DEV restore will replace the lab database.'}
function SqlQuote([string]$s){$s.Replace("'","''")}
function IdQuote([string]$s){'['+$s.Replace(']',']]')+']'}
function Query([string]$sql,[string]$database='master',[int]$timeout=120){
 $cn=[System.Data.SqlClient.SqlConnection]::new("Server=$Server;Database=$database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15")
 try {$cn.Open();$cmd=$cn.CreateCommand();$cmd.CommandText=$sql;$cmd.CommandTimeout=$timeout;$da=[System.Data.SqlClient.SqlDataAdapter]::new($cmd);$dt=[System.Data.DataTable]::new();[void]$da.Fill($dt);return ,$dt} finally {$cn.Dispose()}
}
function Execute([string]$sql,[string]$database='master',[int]$timeout=7200){
 $cn=[System.Data.SqlClient.SqlConnection]::new("Server=$Server;Database=$database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15")
 try {$cn.Open();$cmd=$cn.CreateCommand();$cmd.CommandText=$sql;$cmd.CommandTimeout=$timeout;[void]$cmd.ExecuteNonQuery()}finally{$cn.Dispose()}
}
$targetId=IdQuote $TargetDatabase
$targetLit=SqlQuote $TargetDatabase
$sourceLit=SqlQuote $SourceDatabase
$fullShare=Join-Path $Share 'Full'
$logs='C:\SQLLab\Logs'
New-Item -ItemType Directory -Path $logs -Force | Out-Null
function Write-Audit([string]$stage,[string]$result,[string]$message){
 $out=[pscustomobject]@{Timestamp=(Get-Date).ToString('o');Server=$env:COMPUTERNAME;Stage=$stage;Result=$result;Message=$message}
 $out | ConvertTo-Json -Compress | Add-Content -Path (Join-Path $logs 'dev_restore.jsonl')
}
function Restore-Now {
 $started=Get-Date
 try {
   # Backup from current calendar date only. Never automatically choose yesterday's file.
   $today=(Get-Date).ToString('yyyyMMdd')
   $files=@(Get-ChildItem -LiteralPath $fullShare -Filter "${SourceDatabase}_${today}_*.bak" -File -ErrorAction Stop | Sort-Object LastWriteTime -Descending)
   if($files.Count -eq 0){throw "No full backup for today's date ($today) in $fullShare. Stopping rather than using an old backup."}
   $backup=$files[0].FullName
   $fileEsc=SqlQuote $backup
   Write-Host "Candidate: $backup" -ForegroundColor Cyan
   Execute "RESTORE VERIFYONLY FROM DISK=N'$fileEsc' WITH CHECKSUM;" 'master' 3600
   $header=Query "RESTORE HEADERONLY FROM DISK=N'$fileEsc';" 'master' 180
   if($header.Rows.Count -ne 1 -or [string]$header.Rows[0]['DatabaseName'] -ne $SourceDatabase){throw 'Backup is not an unambiguous full backup of the expected PROD database.'}
   if([int]$header.Rows[0]['BackupType'] -ne 1){throw 'Backup is not a full database backup.'}
   $existing=Query "SELECT name FROM sys.databases WHERE name=N'$targetLit';"
   if($existing.Rows.Count -gt 0){
     $marker=Query "SELECT OBJECT_ID(N'dbo.LabMetadata',N'U') AS MarkerTable" $TargetDatabase
     if($marker.Rows[0]['MarkerTable'] -is [DBNull]){throw 'Target exists without lab marker table. Refusing overwrite.'}
     $mark=Query "SELECT SettingValue FROM dbo.LabMetadata WHERE SettingName=N'DatasetMarker'" $TargetDatabase
     if($mark.Rows.Count -ne 1 -or [string]$mark.Rows[0]['SettingValue'] -ne 'SYNTHETIC-POWERBI-LAB'){throw 'Target not marked synthetic lab; refusing overwrite.'}
   }
   $locations=Query "SELECT CAST(SERVERPROPERTY('InstanceDefaultDataPath') AS nvarchar(4000)) AS DataPath, CAST(SERVERPROPERTY('InstanceDefaultLogPath') AS nvarchar(4000)) AS LogPath;"
   $dataPath=[string]$locations.Rows[0]['DataPath'];$logPath=[string]$locations.Rows[0]['LogPath']
   if([string]::IsNullOrWhiteSpace($dataPath) -or [string]::IsNullOrWhiteSpace($logPath)){throw 'SQL default file directories unavailable.'}
   $fileList=Query "RESTORE FILELISTONLY FROM DISK=N'$fileEsc';"
   $moves=@();$i=0
   foreach($row in $fileList.Rows){
     $i++
     $logical=SqlQuote ([string]$row['LogicalName'])
     $isLog=([string]$row['Type'] -eq 'L')
     $ext=if($isLog){'.ldf'}else{'.mdf'}
     $folder=if($isLog){$logPath}else{$dataPath}
     $physical=Join-Path $folder ("${TargetDatabase}_${i}${ext}")
     $moves+="MOVE N'$logical' TO N'$(SqlQuote $physical)'"
   }
   if($moves.Count -eq 0){throw 'Backup contains no files.'}
   # Snapshot currently existing DEV lab database before any overwrite (extra disk required).
   if($existing.Rows.Count -gt 0){
      $safeDir='C:\SQLLab\PreRestore';New-Item -Path $safeDir -ItemType Directory -Force | Out-Null
      $safeFile=Join-Path $safeDir ($TargetDatabase+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.bak')
      Write-Host "Protecting existing DEV state: $safeFile" -ForegroundColor Yellow
      Execute "BACKUP DATABASE $targetId TO DISK=N'$(SqlQuote $safeFile)' WITH INIT,CHECKSUM,COMPRESSION; RESTORE VERIFYONLY FROM DISK=N'$(SqlQuote $safeFile)' WITH CHECKSUM;" 'master' 3600
   }
   Write-Host 'Starting DEV restore. Active connections to lab DEV database will be disconnected.' -ForegroundColor Yellow
   $restoreSql="IF DB_ID(N'$targetLit') IS NOT NULL ALTER DATABASE $targetId SET SINGLE_USER WITH ROLLBACK IMMEDIATE; `nRESTORE DATABASE $targetId FROM DISK=N'$fileEsc' WITH REPLACE, RECOVERY, " + ($moves -join ', ') + ', STATS=10;'
   try{Execute $restoreSql 'master' 7200}finally{
      try{Execute "IF DB_ID(N'$targetLit') IS NOT NULL AND DATABASEPROPERTYEX(N'$targetLit', 'Status')='ONLINE' ALTER DATABASE $targetId SET MULTI_USER;" 'master' 60}catch{Write-Warning "Unable to set MULTI_USER: $_"}
   }
   $check=Query "SELECT SettingValue FROM dbo.LabMetadata WHERE SettingName=N'DatasetMarker'" $TargetDatabase
   if($check.Rows.Count -ne 1 -or [string]$check.Rows[0]['SettingValue'] -ne 'SYNTHETIC-POWERBI-LAB'){throw 'Post-restore marker check failed.'}
   $counts=Query 'SELECT (SELECT COUNT_BIG(*) FROM dbo.SalesOrders) Orders, (SELECT COUNT_BIG(*) FROM dbo.Payments) Payments;' $TargetDatabase
   $summary="File=$backup Orders=$($counts.Rows[0]['Orders']) Payments=$($counts.Rows[0]['Payments']) DurationSeconds=$([math]::Round(((Get-Date)-$started).TotalSeconds,2))"
   Write-Audit 'Restore' 'SUCCESS' $summary
   Write-Host "DEV restore success. $summary" -ForegroundColor Green
 }catch{Write-Audit 'Restore' 'FAILURE' $_.Exception.Message;throw}
}
if($RunOnce){Restore-Now}
if($InstallJob){
 $expectedPath='C:\SQLLab\02_DEV_Restore.ps1'
 if(-not (Test-Path -LiteralPath $expectedPath)){throw "Copy this script to $expectedPath before installing the job."}
 $job='LAB_DEV_Analytics_Restore_0530';$j=SqlQuote $job
 $command='powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\SQLLab\02_DEV_Restore.ps1" -RunOnce -ApproveOverwrite'
 $cmd=SqlQuote $command
 $tsql=@"
IF NOT EXISTS(SELECT 1 FROM msdb.dbo.sysjobs WHERE name=N'$j')
BEGIN
 EXEC msdb.dbo.sp_add_job @job_name=N'$j',@enabled=0,@description=N'Lab DEV restore from fresh verified PROD backup';
 EXEC msdb.dbo.sp_add_jobstep @job_name=N'$j',@step_name=N'Restore lab DEV',@subsystem=N'CmdExec',@command=N'$cmd',@on_success_action=1,@on_fail_action=2;
 EXEC msdb.dbo.sp_add_jobschedule @job_name=N'$j',@name=N'Every day 05:30',@freq_type=4,@freq_interval=1,@active_start_time=53000;
 EXEC msdb.dbo.sp_add_jobserver @job_name=N'$j';
END;
"@
 Execute $tsql
 Write-Host "SQL Agent job installed: $job. Confirm its owner, CmdExec identity and history in SSMS." -ForegroundColor Green
 Write-Warning 'If backup finishes after 05:30, the restore job fails safely. Configure retry/alerting after validating first runs.'
}
