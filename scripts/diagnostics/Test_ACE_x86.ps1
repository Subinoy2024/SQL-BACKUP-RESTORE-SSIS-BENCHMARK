#requires -Version 5.1
# Run using SysWOW64\WindowsPowerShell\v1.0\powershell.exe (32-bit process).
param([string]$Excel='C:\SQLLab\SSIS\Input\SalesBenchmark_100K.xlsx',[string]$Sheet='SalesData$', [int]$ExpectedRows=100000)
$ErrorActionPreference='Stop'
if([Environment]::Is64BitProcess){throw 'Use 32-bit PowerShell from SysWOW64, not System32.'}
if(-not(Test-Path -LiteralPath $Excel)){throw "Missing workbook: $Excel"}
$cs='Provider=Microsoft.ACE.OLEDB.16.0;Data Source='+$Excel+';Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1";'
$c=[System.Data.OleDb.OleDbConnection]::new($cs)
try{
 $c.Open();$cmd=$c.CreateCommand();$cmd.CommandText='SELECT COUNT(*) FROM ['+$Sheet.Replace(']',']]')+']';$count=[long]$cmd.ExecuteScalar()
 Write-Host "ACE x86 connection: PASSED; rows=$count"
 if($count -ne $ExpectedRows){throw "Expected $ExpectedRows rows, got $count"}
}finally{$c.Dispose()}
