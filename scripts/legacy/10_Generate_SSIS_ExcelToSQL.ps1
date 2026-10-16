#requires -Version 5.1
<#
Creates a genuine SSIS Excel Source -> OLE DB Destination .dtsx via the
SSIS 17.0 managed object model. Run in 64-bit Windows PowerShell 5.1.
Does not write to the SQL benchmark table or execute the package.
#>
[CmdletBinding()]
param(
  [string]$ExcelPath = 'C:\SQLLab\SSIS\Input\SalesBenchmark_100K.xlsx',
  [string]$SheetName = 'SalesData$',
  [string]$SqlServer = '192.168.88.19',
  [string]$Database = 'PowerBIAnalytics_DEV',
  [string]$Destination = '[dbo].[SSIS_Benchmark_Results]',
  [string]$SqlProvider = 'MSOLEDBSQL',
  [string]$PackagePath = 'C:\SQLLab\SSIS\SSIS_Performance_Lab\SalesExcel_100K.dtsx',
  [switch]$OverwritePackage
)
$ErrorActionPreference='Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [IntPtr]::Size -ne 8) { throw 'Run in 64-bit Windows PowerShell 5.1.' }
if (!(Test-Path -LiteralPath $ExcelPath)) { throw "Excel workbook missing: $ExcelPath" }
if ((Test-Path -LiteralPath $PackagePath) -and !$OverwritePackage) { throw "Package exists: $PackagePath. Use -OverwritePackage only if replacement is intended." }
$assemblyRoot='C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\PublicAssemblies\SSIS\170'
foreach($n in @('Microsoft.SqlServer.DTSRuntimeWrap.dll','Microsoft.SqlServer.DTSPipelineWrap.dll','Microsoft.SqlServer.ManagedDTS.dll')) {
  $p=Join-Path $assemblyRoot $n
  if (!(Test-Path -LiteralPath $p)) { throw "Missing: $p" }
  [void][Reflection.Assembly]::LoadFrom($p)
}
# Check that the process can actually open the intended workbook/worksheet.
$ace = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$ExcelPath;Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=0`";"
$testExcel = New-Object System.Data.OleDb.OleDbConnection($ace)
try {
  $testExcel.Open()
  $cmd = $testExcel.CreateCommand()
  $escapedSheet=$SheetName.Replace(']',']]')
  $cmd.CommandText="SELECT TOP 1 * FROM [$escapedSheet]"
  $reader=$cmd.ExecuteReader()
  try {
    $names=@(0..($reader.FieldCount-1) | ForEach-Object {$reader.GetName($_)})
    foreach($required in @('OrderID','ProductID','Quantity','SalesAmount')) {
      if ($names -notcontains $required) { throw "Excel column missing: $required; found: $($names -join ', ')" }
    }
  } finally { $reader.Close() }
} finally { $testExcel.Dispose() }
# Check OLE DB provider and destination metadata. Does not insert/delete rows.
$sqlCS="Provider=$SqlProvider;Data Source=$SqlServer;Initial Catalog=$Database;Integrated Security=SSPI;"
$testSql=New-Object System.Data.OleDb.OleDbConnection($sqlCS)
try {
  $testSql.Open()
  $cmd=$testSql.CreateCommand(); $cmd.CommandText="SELECT TOP 0 OrderID, ProductID, Quantity, SalesAmount FROM $Destination"
  $reader=$cmd.ExecuteReader();$reader.Close()
} catch {
  throw "SQL OLE DB test failed (Provider=$SqlProvider). Install a compatible x64 MSOLEDBSQL driver or pass -SqlProvider. Details: $($_.Exception.Message)"
} finally { $testSql.Dispose() }

$app=New-Object Microsoft.SqlServer.Dts.Runtime.Application
$pkg=New-Object Microsoft.SqlServer.Dts.Runtime.Package
$pkg.Name='SalesExcel_100K'
$pkg.Description='100K Excel Source to SQL Server benchmark. Created using the SSIS 17.0 object model.'
$pkg.ProtectionLevel=[Microsoft.SqlServer.Dts.Runtime.DTSProtectionLevel]::DontSaveSensitive

$excelCM=$pkg.Connections.Add('EXCEL')
$excelCM.Name='Benchmark Excel'
$excelCM.ConnectionString=$ace
$sqlCM=$pkg.Connections.Add('OLEDB')
$sqlCM.Name='Benchmark SQL'
$sqlCM.ConnectionString=$sqlCS

$task=[Microsoft.SqlServer.Dts.Runtime.TaskHost]$pkg.Executables.Add('STOCK:PipelineTask')
$task.Name='Excel to SQL 100K'
$pipe=[Microsoft.SqlServer.Dts.Pipeline.Wrapper.MainPipe]$task.InnerObject

$source=$pipe.ComponentMetaDataCollection.New()
$source.Name='Excel Source'
$source.ComponentClassID='DTSAdapter.ExcelSource'
$src=$source.Instantiate()
$src.ProvideComponentProperties()
$source.RuntimeConnectionCollection[0].ConnectionManager=[Microsoft.SqlServer.Dts.Runtime.DtsConvert]::GetExtendedInterface($excelCM)
$source.RuntimeConnectionCollection[0].ConnectionManagerID=$excelCM.ID
$src.SetComponentProperty('AccessMode',0)
$src.SetComponentProperty('OpenRowset',"[$escapedSheet]")
$src.AcquireConnections($null)
try { $src.ReinitializeMetaData() } finally { $src.ReleaseConnections() }

$destinationComponent=$pipe.ComponentMetaDataCollection.New()
$destinationComponent.Name='OLE DB Destination'
$destinationComponent.ComponentClassID='DTSAdapter.OleDbDestination'
$dst=$destinationComponent.Instantiate()
$dst.ProvideComponentProperties()
$destinationComponent.RuntimeConnectionCollection[0].ConnectionManager=[Microsoft.SqlServer.Dts.Runtime.DtsConvert]::GetExtendedInterface($sqlCM)
$destinationComponent.RuntimeConnectionCollection[0].ConnectionManagerID=$sqlCM.ID
# AccessMode=3 denotes fast load into an OLE DB table.
$dst.SetComponentProperty('AccessMode',3)
$dst.SetComponentProperty('OpenRowset',$Destination)
$dst.AcquireConnections($null)
try { $dst.ReinitializeMetaData() } finally { $dst.ReleaseConnections() }

$path=$pipe.PathCollection.New()
$path.AttachPathAndPropagateNotifications($source.OutputCollection[0],$destinationComponent.InputCollection[0])
$input=$destinationComponent.InputCollection[0]
$virtual=$input.GetVirtualInput()
$wanted=@('OrderID','ProductID','Quantity','SalesAmount')
foreach($columnName in $wanted) {
  $candidate=$null
  foreach($v in $virtual.VirtualInputColumnCollection) { if ($v.Name -eq $columnName) { $candidate=$v;break } }
  if (!$candidate) { throw "SSIS source virtual input missing: $columnName" }
  $added=$dst.SetUsageType($input.ID,$virtual,$candidate.LineageID,[Microsoft.SqlServer.Dts.Pipeline.Wrapper.DTSUsageType]::UT_READONLY)
  $external=$null
  foreach($e in $input.ExternalMetadataColumnCollection) { if ($e.Name -eq $columnName) { $external=$e;break } }
  if (!$external) { throw "SSIS destination external metadata missing: $columnName" }
  $dst.MapInputColumn($input.ID,$added.ID,$external.ID)
}
# Validate package metadata without running the pipeline.
$validation=$pkg.Validate($null,$null,$null,$null)
Write-Host "Package validation result: $validation"
if ($validation -ne [Microsoft.SqlServer.Dts.Runtime.DTSExecResult]::Success) {
  throw "SSIS package validation failed ($validation). No package saved; inspect provider, metadata and SSIS installation."
}
$folder=Split-Path -Parent $PackagePath
New-Item -Path $folder -ItemType Directory -Force | Out-Null
$app.SaveToXml($PackagePath,$pkg,$null)
Write-Host "Generated SSIS package: $PackagePath" -ForegroundColor Green
Write-Host 'No data was loaded. Run DTExec /Validate before the benchmark.'
