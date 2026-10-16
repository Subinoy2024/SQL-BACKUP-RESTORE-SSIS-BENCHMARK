# Troubleshooting seen in the original lab

| Error / symptom | Cause discovered or likely issue | What worked / safe next step |
|---|---|---|
| `Cannot convert System.__ComObject to MainPipe` | PowerShell dynamic COM cast | Use a compiled, typed C# helper rather than casting pipeline COM wrappers from PowerShell |
| `0x8002801D TYPE_E_LIBNOTREGISTERED` at `ProvideComponentProperties` | PowerShell/COM interoperability path | Typed C# successfully initialized Excel Source, showing it was not necessarily a broken SSIS installation |
| `Invalid object name dbo.SSIS_Benchmark_Results` | Daily DEV restore removed the table | Recreate the empty lab table (never a business table) and ensure SQL connection targets the correct database |
| `0xC02020E8` on Excel rowset | Sheet name/Excel metadata issue | Verify `SalesData$` exists and is referred to consistently; inspect Excel table metadata |
| `0xC0048021` Excel Source initialization | Unversioned component ProgID did not work | Resolve installed component ProgIDs; the working 2025 lab used `DTSAdapter.ExcelSource.9` / `DTSAdapter.OleDbDestination.9` |
| `TYPE_ALREADY_EXISTS` with `Add-Type` | Reloading a compiled class in the same PowerShell process | Run script in a fresh 64-bit PowerShell 5.1 process |
| `Microsoft.ACE.OLEDB.16.0 provider is not registered` in x86 | Only x64 Access Runtime installed; x86 COM DLL absent | Confirm x86 COM in-process DLL and ACE connection in 32-bit PowerShell; plan supported x86 runtime migration separately |
| Missing `setup.exe` for Office Deployment Tool | ODT installer not extracted | Download official ODT and extract its `setup.exe` first |
| SSIS validates but no data loaded | `/Validate` checks metadata only | Run package explicitly, check native exit code and 100,000 destination rows |

Read the SSIS DTExec output and logged exception stage before changing SQL or Windows registry. Avoid copying DLLs across architectures or using `regsvr32` as a guess.
