#requires -Version 5.1
<#
 Run in elevated Windows PowerShell 5.1 on 192.168.88.18.
 Creates a synthetic Power BI lab database, readonly reporting login via Windows group (manual),
 local backup folder, and 05:00 SQL Agent backup job. Never replaces existing database.
 #>
[CmdletBinding()]
param(
 [string]$Server='localhost',
 [string]$Database='PowerBIAnalytics_PROD',
 [string]$BackupRoot='C:\DBBACKUP',
 [int]$CustomerCount=1000,
 [int]$ProductCount=100,
 [int]$OrderCount=20000,
 [switch]$RunBackupNow
)
$ErrorActionPreference='Stop'
if($Database -ne 'PowerBIAnalytics_PROD'){throw 'This lab script is restricted to PowerBIAnalytics_PROD.'}
if($CustomerCount -lt 10 -or $ProductCount -lt 10 -or $OrderCount -lt 100){throw 'Data sizes below minimum.'}
function Invoke-DbSql {
 param([string]$Sql,[string]$Db='master',[int]$Timeout=180)
 $c=[System.Data.SqlClient.SqlConnection]::new("Server=$Server;Database=$Db;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15")
 try { $c.Open(); $q=$c.CreateCommand();$q.CommandText=$Sql;$q.CommandTimeout=$Timeout;$null=$q.ExecuteNonQuery() } finally { $c.Dispose() }
}
function Quote-Sql([string]$s){return $s.Replace("'","''")}
function Quote-Ident([string]$s){return '['+$s.Replace(']',']]')+']'}
$DbIdent=Quote-Ident $Database
$DbLit=Quote-Sql $Database
$FullDir=Join-Path $BackupRoot 'Full'
New-Item -Path $FullDir -ItemType Directory -Force | Out-Null
Write-Host "Using $FullDir. Review inherited NTFS permissions; do not grant Everyone write access." -ForegroundColor Yellow
$create=@"
IF DB_ID(N'$DbLit') IS NULL BEGIN
 CREATE DATABASE $DbIdent;
END ELSE IF EXISTS (SELECT 1 FROM $DbIdent.sys.tables)
 AND NOT EXISTS (SELECT 1 FROM $DbIdent.sys.tables WHERE name = N'LabMetadata') BEGIN
 THROW 51000, 'Existing database contains non-lab tables. Refusing changes.', 1;
END;
"@
Invoke-DbSql $create
$schema=@'
IF OBJECT_ID('dbo.Customers','U') IS NULL
 CREATE TABLE dbo.Customers(CustomerID int NOT NULL PRIMARY KEY,CustomerName nvarchar(140) NOT NULL,Region nvarchar(40) NOT NULL,JoinDate date NOT NULL);
IF OBJECT_ID('dbo.Products','U') IS NULL
 CREATE TABLE dbo.Products(ProductID int NOT NULL PRIMARY KEY,ProductName nvarchar(140) NOT NULL,Category nvarchar(40) NOT NULL,UnitPrice decimal(12,2) NOT NULL);
IF OBJECT_ID('dbo.SalesOrders','U') IS NULL
 CREATE TABLE dbo.SalesOrders(OrderID int NOT NULL PRIMARY KEY,CustomerID int NOT NULL REFERENCES dbo.Customers(CustomerID),OrderDate datetime2(0) NOT NULL,OrderStatus nvarchar(20) NOT NULL);
IF OBJECT_ID('dbo.SalesOrderDetails','U') IS NULL
 CREATE TABLE dbo.SalesOrderDetails(OrderID int NOT NULL REFERENCES dbo.SalesOrders(OrderID),[LineNo] tinyint NOT NULL,ProductID int NOT NULL REFERENCES dbo.Products(ProductID),Quantity int NOT NULL,UnitPrice decimal(12,2) NOT NULL,CONSTRAINT PK_SalesOrderDetails PRIMARY KEY(OrderID,[LineNo]));
IF OBJECT_ID('dbo.Payments','U') IS NULL
 CREATE TABLE dbo.Payments(PaymentID int NOT NULL PRIMARY KEY,OrderID int NOT NULL UNIQUE REFERENCES dbo.SalesOrders(OrderID),PaymentDate datetime2(0) NOT NULL,Amount decimal(14,2) NOT NULL,PaymentMethod nvarchar(30) NOT NULL);
IF OBJECT_ID('dbo.SalesTargets','U') IS NULL
 CREATE TABLE dbo.SalesTargets(TargetMonth date NOT NULL,Region nvarchar(40) NOT NULL,TargetAmount decimal(14,2) NOT NULL,CONSTRAINT PK_SalesTargets PRIMARY KEY(TargetMonth,Region));
IF OBJECT_ID('dbo.LabMetadata','U') IS NULL CREATE TABLE dbo.LabMetadata(SettingName sysname PRIMARY KEY,SettingValue nvarchar(200) NOT NULL);
IF NOT EXISTS(SELECT 1 FROM dbo.LabMetadata WHERE SettingName=N'DatasetMarker')
 INSERT dbo.LabMetadata VALUES(N'DatasetMarker',N'SYNTHETIC-POWERBI-LAB');
'@
Invoke-DbSql $schema $Database
$seed=@"
SET NOCOUNT ON;
IF NOT EXISTS (SELECT 1 FROM dbo.Customers)
BEGIN
 ;WITH n AS (SELECT TOP ($CustomerCount) ROW_NUMBER() OVER(ORDER BY (SELECT NULL)) n FROM sys.all_objects a CROSS JOIN sys.all_objects b)
 INSERT dbo.Customers SELECT n, CONCAT(N'Customer ',RIGHT(CONCAT('000000',n),6)),CHOOSE((n%4)+1,N'North',N'South',N'East',N'West'),DATEADD(DAY,-(n%900),CONVERT(date,GETDATE())) FROM n;
END;
IF NOT EXISTS (SELECT 1 FROM dbo.Products)
BEGIN
 ;WITH n AS (SELECT TOP ($ProductCount) ROW_NUMBER() OVER(ORDER BY (SELECT NULL)) n FROM sys.all_objects a CROSS JOIN sys.all_objects b)
 INSERT dbo.Products SELECT n,CONCAT(N'Product ',n),CHOOSE((n%5)+1,N'Hardware',N'Software',N'Services',N'Supplies',N'Accessories'),CAST((20+(n%200)*3.25) AS decimal(12,2)) FROM n;
END;
IF NOT EXISTS (SELECT 1 FROM dbo.SalesOrders)
BEGIN
 ;WITH n AS (SELECT TOP ($OrderCount) ROW_NUMBER() OVER(ORDER BY (SELECT NULL)) n FROM sys.all_objects a CROSS JOIN sys.all_objects b)
 INSERT dbo.SalesOrders SELECT n,((n-1)%$CustomerCount)+1,DATEADD(HOUR,-(n%8760),SYSUTCDATETIME()),CASE WHEN n%20=0 THEN N'Cancelled' ELSE N'Completed' END FROM n;
 INSERT dbo.SalesOrderDetails(OrderID,[LineNo],ProductID,Quantity,UnitPrice)
 SELECT o.OrderID,1,((o.OrderID-1)%$ProductCount)+1,(o.OrderID%5)+1,p.UnitPrice
 FROM dbo.SalesOrders o JOIN dbo.Products p ON p.ProductID=((o.OrderID-1)%$ProductCount)+1;
 INSERT dbo.Payments(PaymentID,OrderID,PaymentDate,Amount,PaymentMethod)
 SELECT o.OrderID,o.OrderID,DATEADD(MINUTE,30,o.OrderDate),CAST(d.Quantity*d.UnitPrice AS decimal(14,2)),CHOOSE((o.OrderID%3)+1,N'Card',N'UPI',N'Bank')
 FROM dbo.SalesOrders o JOIN dbo.SalesOrderDetails d ON o.OrderID=d.OrderID
 WHERE o.OrderStatus=N'Completed';
END;
IF NOT EXISTS (SELECT 1 FROM dbo.SalesTargets)
BEGIN
 ;WITH n AS (SELECT 0 n UNION ALL SELECT n+1 FROM n WHERE n<11)
 INSERT dbo.SalesTargets SELECT DATEADD(MONTH,-n,DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1)),r.Region,CAST(100000 AS decimal(14,2))
 FROM n CROSS JOIN (VALUES(N'North'),(N'South'),(N'East'),(N'West')) r(Region) OPTION (MAXRECURSION 100);
END;
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE name='IX_SalesOrders_OrderDate' AND object_id=OBJECT_ID('dbo.SalesOrders'))
 CREATE INDEX IX_SalesOrders_OrderDate ON dbo.SalesOrders(OrderDate) INCLUDE(CustomerID,OrderStatus);
IF OBJECT_ID('dbo.vw_SalesAnalytics','V') IS NULL EXEC(N'CREATE VIEW dbo.vw_SalesAnalytics AS SELECT 1 AS Placeholder');
EXEC(N'ALTER VIEW dbo.vw_SalesAnalytics AS
 SELECT o.OrderID,o.OrderDate,o.OrderStatus,c.CustomerID,c.CustomerName,c.Region,
 p.ProductID,p.ProductName,p.Category,d.Quantity,d.UnitPrice,
 CAST(d.Quantity*d.UnitPrice AS decimal(14,2)) AS SalesAmount
 FROM dbo.SalesOrders o JOIN dbo.Customers c ON c.CustomerID=o.CustomerID
 JOIN dbo.SalesOrderDetails d ON d.OrderID=o.OrderID
 JOIN dbo.Products p ON p.ProductID=d.ProductID');
"@
Invoke-DbSql $seed $Database 600
$backupFile=Join-Path $FullDir ($Database+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.bak')
$SqlBackupFile=Quote-Sql $backupFile
$backupQuery="BACKUP DATABASE $DbIdent TO DISK=N'$SqlBackupFile' WITH INIT, CHECKSUM, COMPRESSION, STATS=10; RESTORE VERIFYONLY FROM DISK=N'$SqlBackupFile' WITH CHECKSUM;"
# Validate SQL Agent availability before creating a job.
$jobName='LAB_PROD_Analytics_Backup_0500'
$jobEsc=Quote-Sql $jobName
# Job generates unique filenames and verifies backup; it does not delete old backups.
$backupJobCommand=@"
DECLARE @path nvarchar(4000)=N'$(Quote-Sql $FullDir)\$Database'+'_'+CONVERT(char(8),GETDATE(),112)+'_'+REPLACE(CONVERT(char(8),GETDATE(),108),':','')+'.bak';
BACKUP DATABASE $DbIdent TO DISK=@path WITH INIT,CHECKSUM,COMPRESSION,STATS=10;
RESTORE VERIFYONLY FROM DISK=@path WITH CHECKSUM;
"@
$escapedCmd=Quote-Sql $backupJobCommand
$jobSql=@"
IF NOT EXISTS(SELECT 1 FROM msdb.dbo.sysjobs WHERE name=N'$jobEsc')
BEGIN
 EXEC msdb.dbo.sp_add_job @job_name=N'$jobEsc',@enabled=0,@description=N'Lab full backup and VERIFYONLY, PROD SQL Server';
 EXEC msdb.dbo.sp_add_jobstep @job_name=N'$jobEsc',@step_name=N'Backup and verify',@subsystem=N'TSQL',@database_name=N'master',@command=N'$escapedCmd',@on_success_action=1,@on_fail_action=2;
 EXEC msdb.dbo.sp_add_jobschedule @job_name=N'$jobEsc',@name=N'Every day 05:00',@freq_type=4,@freq_interval=1,@active_start_time=50000;
 EXEC msdb.dbo.sp_add_jobserver @job_name=N'$jobEsc';
END;
"@
Invoke-DbSql $jobSql
if($RunBackupNow){Write-Host 'Creating verified full backup...' -ForegroundColor Cyan;Invoke-DbSql $backupQuery 'master' 3600;Write-Host "Created: $backupFile" -ForegroundColor Green}
Write-Host "PROD ready. DB: $Database; backup folder: $FullDir; SQL Agent job: $jobName" -ForegroundColor Green
Write-Host 'Power BI Desktop: Get Data > SQL Server > 192.168.88.18 > PowerBIAnalytics_PROD > dbo.vw_SalesAnalytics (Windows authentication).' -ForegroundColor Cyan
Write-Warning 'Power BI connection permissions must be granted to a dedicated read-only Windows account/group; do not distribute vmadmin credentials.'
Write-Warning 'Add backup retention, independent off-host backup, SQL Agent notifications, and disk alerts before production use.'
