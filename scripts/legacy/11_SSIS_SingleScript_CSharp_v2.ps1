#requires -Version 5.1
<#
Single-file SSIS 2025 Excel -> SQL generator using compiled C# for typed COM interop.
Default: generate & validate only. -RunBenchmark -AllowResetBenchmarkTable: run x64 three times.
No PROD changes. Validated on the user's hardware only after successful DTExec /Validate.
#>
[CmdletBinding()]
param(
 [string]$ExcelPath='C:\SQLLab\SSIS\Input\SalesBenchmark_100K.xlsx',
 [string]$SheetName='SalesData$',
 [string]$SqlServer='192.168.88.19',
 [string]$Database='PowerBIAnalytics_DEV',
 [string]$SqlProvider='MSOLEDBSQL',
 [string]$PackagePath='C:\SQLLab\SSIS\SSIS_Performance_Lab\SalesExcel_100K.dtsx',
 [int]$ExpectedRows=100000,
 [ValidateRange(1,20)][int]$Runs=3,
 [switch]$OverwritePackage,
 [switch]$RunBenchmark,
 [switch]$AllowResetBenchmarkTable
)
$ErrorActionPreference='Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [IntPtr]::Size -ne 8) { throw 'Run 64-bit Windows PowerShell 5.1 (not PowerShell 7).' }
if ($Database -ne 'PowerBIAnalytics_DEV') { throw 'Safety stop: database must be PowerBIAnalytics_DEV.' }
if (!(Test-Path -LiteralPath $ExcelPath -PathType Leaf)) { throw "Excel workbook missing: $ExcelPath" }
if ((Test-Path -LiteralPath $PackagePath) -and !$OverwritePackage) { throw "Package already exists; use -OverwritePackage to replace: $PackagePath" }
$dtExec='C:\Program Files\Microsoft SQL Server\170\DTS\Binn\DTExec.exe'
if (!(Test-Path -LiteralPath $dtExec)) { throw "Missing DTExec: $dtExec" }
$root='C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\PublicAssemblies\SSIS\170'
$dlls=@('Microsoft.SqlServer.DTSRuntimeWrap.dll','Microsoft.SqlServer.DTSPipelineWrap.dll','Microsoft.SqlServer.ManagedDTS.dll') | ForEach-Object { Join-Path $root $_ }
foreach($d in $dlls) { if (!(Test-Path -LiteralPath $d)) { throw "SSIS assembly missing: $d" }; [void][Reflection.Assembly]::LoadFrom($d) }
$excelCS="Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$ExcelPath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=0`";"
$sqlCS="Provider=$SqlProvider;Data Source=$SqlServer;Initial Catalog=$Database;Integrated Security=SSPI;TrustServerCertificate=True;"
# Read-only source and target checks.
$excelConn=[System.Data.OleDb.OleDbConnection]::new($excelCS)
try {
 $excelConn.Open()
 $cmd=$excelConn.CreateCommand(); $cmd.CommandText='SELECT TOP 1 * FROM [' + $SheetName.Replace(']',']]') + ']'
 $reader=$cmd.ExecuteReader()
 try { $names=@(for($n=0;$n -lt $reader.FieldCount;$n++) {$reader.GetName($n)}) } finally { $reader.Close() }
 foreach($n in @('OrderID','ProductID','Quantity','SalesAmount')) { if($names -notcontains $n) { throw "Excel column missing: $n" } }
} finally { $excelConn.Dispose() }
$sqlConn=[System.Data.OleDb.OleDbConnection]::new($sqlCS)
try {
 $sqlConn.Open()
 $cmd=$sqlConn.CreateCommand(); $cmd.CommandText='SELECT TOP 0 OrderID, ProductID, Quantity, SalesAmount FROM [dbo].[SSIS_Benchmark_Results]'
 $reader=$cmd.ExecuteReader(); $reader.Close()
} finally { $sqlConn.Dispose() }

# Typed C# interop avoids PowerShell's 0x8002801D at ProvideComponentProperties().
$cs=@'
using System;
using System.IO;
using Microsoft.SqlServer.Dts.Runtime;
using Microsoft.SqlServer.Dts.Pipeline.Wrapper;

public static class SSIS2025Generator
{
    public static void Generate(string excelCS, string sheet, string sqlCS, string destination, string packagePath)
    {
        var app = new Application();
        var pkg = new Package();
        pkg.Name = "SalesExcel_100K";
        pkg.Description = "Genuine Excel Source to OLE DB Destination benchmark generated with typed SSIS 17 interop";
        pkg.ProtectionLevel = DTSProtectionLevel.DontSaveSensitive;

        ConnectionManager excelCM = pkg.Connections.Add("EXCEL");
        excelCM.Name = "Benchmark Excel";
        excelCM.ConnectionString = excelCS;
        ConnectionManager sqlCM = pkg.Connections.Add("OLEDB");
        sqlCM.Name = "Benchmark SQL";
        sqlCM.ConnectionString = sqlCS;

        TaskHost host = (TaskHost)pkg.Executables.Add("STOCK:PipelineTask");
        host.Name = "Excel to SQL 100K";
        MainPipe pipe = (MainPipe)host.InnerObject;

        IDTSComponentMetaData100 srcMeta = pipe.ComponentMetaDataCollection.New();
        srcMeta.Name = "Excel Source";
        srcMeta.ComponentClassID = "DTSAdapter.ExcelSource";
        CManagedComponentWrapper src = srcMeta.Instantiate();
        src.ProvideComponentProperties();
        srcMeta.RuntimeConnectionCollection[0].ConnectionManager = DtsConvert.GetExtendedInterface(excelCM);
        srcMeta.RuntimeConnectionCollection[0].ConnectionManagerID = excelCM.ID;
        src.SetComponentProperty("AccessMode", 0);
        src.SetComponentProperty("OpenRowset", sheet);
        Console.WriteLine("STAGE: Excel Source AcquireConnections");
        src.AcquireConnections(null);
        try { Console.WriteLine("STAGE: Excel Source ReinitializeMetaData; worksheet=" + sheet); src.ReinitializeMetaData(); } finally { src.ReleaseConnections(); }

        IDTSComponentMetaData100 dstMeta = pipe.ComponentMetaDataCollection.New();
        dstMeta.Name = "OLE DB Destination";
        dstMeta.ComponentClassID = "DTSAdapter.OleDbDestination";
        CManagedComponentWrapper dst = dstMeta.Instantiate();
        dst.ProvideComponentProperties();
        dstMeta.RuntimeConnectionCollection[0].ConnectionManager = DtsConvert.GetExtendedInterface(sqlCM);
        dstMeta.RuntimeConnectionCollection[0].ConnectionManagerID = sqlCM.ID;
        dst.SetComponentProperty("AccessMode", 3); // fast-load to table
        dst.SetComponentProperty("OpenRowset", destination);
        Console.WriteLine("STAGE: SQL Destination AcquireConnections");
        dst.AcquireConnections(null);
        try { Console.WriteLine("STAGE: SQL Destination ReinitializeMetaData"); dst.ReinitializeMetaData(); } finally { dst.ReleaseConnections(); }

        Console.WriteLine("STAGE: Connecting SSIS pipeline");
        IDTSPath100 path = pipe.PathCollection.New();
        path.AttachPathAndPropagateNotifications(srcMeta.OutputCollection[0], dstMeta.InputCollection[0]);
        IDTSInput100 input = dstMeta.InputCollection[0];
        IDTSVirtualInput100 virtualInput = input.GetVirtualInput();
        foreach (string name in new [] { "OrderID", "ProductID", "Quantity", "SalesAmount" })
        {
            IDTSVirtualInputColumn100 found = null;
            foreach (IDTSVirtualInputColumn100 candidate in virtualInput.VirtualInputColumnCollection)
                if (string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase)) { found = candidate; break; }
            if (found == null) throw new Exception("Missing source column: " + name);
            IDTSInputColumn100 added = dst.SetUsageType(input.ID, virtualInput, found.LineageID, DTSUsageType.UT_READONLY);
            IDTSExternalMetadataColumn100 external = null;
            foreach (IDTSExternalMetadataColumn100 item in input.ExternalMetadataColumnCollection)
                if (string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)) { external = item; break; }
            if (external == null) throw new Exception("Missing destination metadata column: " + name);
            dst.MapInputColumn(input.ID, added.ID, external.ID);
        }
        string folder = Path.GetDirectoryName(packagePath);
        Directory.CreateDirectory(folder);
        Console.WriteLine("STAGE: Saving package");
        app.SaveToXml(packagePath, pkg, null);
    }
}
'@
Write-Host "`n=== COMPILE TYPED SSIS GENERATOR ===" -ForegroundColor Cyan
Add-Type -TypeDefinition $cs -ReferencedAssemblies $dlls -ErrorAction Stop
Write-Host "`n=== GENERATE PACKAGE ===" -ForegroundColor Cyan
try {
 [SSIS2025Generator]::Generate($excelCS,$SheetName,$sqlCS,'[dbo].[SSIS_Benchmark_Results]',$PackagePath)
} catch {
 Write-Host ('Generator exception: ' + $_.Exception.ToString()) -ForegroundColor Red
 throw
}
if (!(Test-Path -LiteralPath $PackagePath)) { throw 'Package not saved.' }
Write-Host "Generated $PackagePath" -ForegroundColor Green
Write-Host "`n=== DTEXEC VALIDATION (NO ROW LOAD) ===" -ForegroundColor Cyan
& $dtExec /F $PackagePath /Validate /REPORTING E
if ($LASTEXITCODE -ne 0) { throw "DTExec /Validate failed, exit code $LASTEXITCODE. Do not run benchmark." }
Write-Host 'DTExec validation passed.' -ForegroundColor Green
if (!$RunBenchmark) { Write-Host 'Generation only. To benchmark, rerun with -OverwritePackage -RunBenchmark -AllowResetBenchmarkTable'; return }
if (!$AllowResetBenchmarkTable) { throw 'Benchmark requires -AllowResetBenchmarkTable (truncates ONLY dbo.SSIS_Benchmark_Results before each run).' }
if ((Get-Date).TimeOfDay -ge [TimeSpan]::FromHours(4.5) -and (Get-Date).TimeOfDay -le [TimeSpan]::FromHours(6.5)) { throw 'Do not benchmark during the 04:30-06:30 DEV restore window.' }
$results='C:\SQLLab\SSIS\Results'
New-Item -ItemType Directory -Force -Path $results | Out-Null
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$rows=@()
$adoCS="Server=$SqlServer;Database=$Database;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;"
for($i=1;$i -le $Runs;$i++) {
 $db=[System.Data.SqlClient.SqlConnection]::new($adoCS)
 try {
  $db.Open(); $cmd=$db.CreateCommand();$cmd.CommandText='TRUNCATE TABLE dbo.SSIS_Benchmark_Results'; [void]$cmd.ExecuteNonQuery()
 } finally { $db.Dispose() }
 $log=Join-Path $results "SSIS_x64_${stamp}_$i.log"
 $timer=[System.Diagnostics.Stopwatch]::StartNew()
 & $dtExec /F $PackagePath /REPORTING E *>&1 | Out-File -LiteralPath $log -Encoding UTF8
 $exit=$LASTEXITCODE; $timer.Stop()
 $db=[System.Data.SqlClient.SqlConnection]::new($adoCS)
 try { $db.Open();$cmd=$db.CreateCommand();$cmd.CommandText='SELECT COUNT_BIG(*) FROM dbo.SSIS_Benchmark_Results';$count=[long]$cmd.ExecuteScalar() } finally { $db.Dispose() }
 $pass=($exit -eq 0 -and $count -eq $ExpectedRows)
 $rows += [pscustomobject]@{Architecture='x64';Run=$i;Seconds=[math]::Round($timer.Elapsed.TotalSeconds,3);ExitCode=$exit;Rows=$count;Passed=$pass;Log=$log}
 Write-Host "Run $i`: seconds=$([math]::Round($timer.Elapsed.TotalSeconds,3)) rows=$count exit=$exit pass=$pass"
 if (!$pass) { break }
}
$csv=Join-Path $results "SSIS_x64_${stamp}.csv"
$rows | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $csv
Write-Host "Results: $csv" -ForegroundColor Cyan
if (@($rows | Where-Object { -not $_.Passed }).Count) { throw 'Benchmark failed. Inspect the run log; no timings should be treated as successful.' }
