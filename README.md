# SQL Backup, Restore, Power BI & SSIS Benchmark Lab

A practical, reusable **Windows SQL Server data engineering lab**. It demonstrates a synthetic sales database, scheduled backup, DEV refresh, Power BI Desktop reporting, and a real SSIS Excel-to-SQL benchmark measured in both 64-bit and 32-bit runtimes.

> **Status:** The original UAT run succeeded. The scripts here are the project sources and a reusable starting point, **not** a turnkey production platform. Always review paths, identities, database names, scheduling and destructive operations before adapting them.

## What you will learn

1. Create a small SQL Server reporting dataset without customer data.
2. Schedule verified, checksummed backups with SQL Server Agent.
3. Restore PROD lab backups to a separate DEV database over a network share.
4. Build a Power BI Desktop report from a SQL view.
5. Generate a genuine SSIS `.dtsx` package in PowerShell using typed C# SSIS APIs (no manual drag-and-drop required for this tested setup).
6. Measure Excel-to-SQL load execution under x64 and x86 SSIS, checking the row count after every run.

## Architecture

```text
[PROD SQL Server]                                   [DEV / UAT SQL Server]
 PowerBIAnalytics_PROD                               PowerBIAnalytics_DEV
          |                                                   ^
          v                                                   |
 SQL Agent 05:00 -> compressed CHECKSUM .bak -> SMB -> SQL Agent 05:30 restore
          |                                                   |
          +--> dbo.vw_SalesAnalytics --> Power BI Desktop     |
                                                              |
 Excel workbook (100,000 synthetic rows)                      |
     -> SSIS Excel Source -> Data Flow -> OLE DB Destination --+
           ^ x64 or x86 DTExec (separate ACE architectures tested sequentially)
```

## Verified lab outcome (2026-10-16)

| Check | Outcome |
|---|---|
| PROD lab dataset | 1,000 customers, 100 products, 20,000 orders, 19,000 payments, 20,000 detail rows |
| PROD backup + DEV restore | Manual Agent test succeeded; both jobs were installed/enabled |
| Power BI Desktop | Sales dashboard with seven visuals rendered after data refresh |
| SSIS `.dtsx` | Generated using typed C# and version-specific source/destination ProgIDs (`.9`); `DTExec /Validate` succeeded |
| x64 load | 3/3 runs, 100,000 rows per run; **3.524 s median** |
| x86 load | 3/3 runs, 100,000 rows per run; **3.991 s median** |

For this small UAT workload, x86 median elapsed time was **13.25% longer** than x64. This does **not** prove a universal 64-bit advantage: three repetitions, different ACE installations, sequential architecture migration and variable caches limit inference. See [Benchmark Method](docs/BENCHMARK.md).

## Repository map

| Folder | Purpose |
|---|---|
| `scripts/prod` | Synthetic analytics database and scheduled backup |
| `scripts/dev` | Date-scoped backup discovery, validation, restore and scheduled job |
| `scripts/powerbi` | Power BI project/report generator used during the project |
| `scripts/ssis` | **Working v3** SSIS generator and architecture-specific runner |
| `scripts/diagnostics` | Read-only host/driver checks |
| `scripts/legacy` | Earlier experiments and superseded scripts kept for learning; **do not run as the primary workflow** |
| `docs` | Beginner-friendly, step-by-step guides and limitations |
| `results` | Example results only; actual UAT logs/CSVs remain on the DEV server |

## Quick start

**Prerequisites:** Windows PowerShell 5.1; SQL Server 2025 Database Engine and Agent; two separate authorized lab SQL instances; domain/SQL rights; sufficient disk space; SSIS 17.0, ACE x64 or x86 as applicable; Visual Studio 2026 with compatible SSIS 170 assemblies **for generating a new DTSX**; Power BI Desktop for the report. The generator's locations and safety checks are intentionally pinned to the originally tested setup and must be reviewed for other environments.

1. **Read** [Full Implementation Guide](docs/IMPLEMENTATION_GUIDE.md) and [Safety](docs/SAFETY.md).
2. On PROD, review `scripts/prod/01_PROD_Setup_Backup.ps1` and run it with `-RunBackupNow` after configuring SQL permissions.
3. On DEV, adapt the backup share and paths in `scripts/dev/02_DEV_Restore.ps1`; run with `-RunOnce -ApproveOverwrite` **only for the synthetic lab database**. Do not point it at a real DEV database.
4. On a Windows machine with Power BI Desktop, run `scripts/powerbi/07_EndToEnd_PowerBI.ps1`, open the generated `.pbip`, enter credentials, refresh, check the visuals.
5. Create a **synthetic** `SalesBenchmark_100K.xlsx` with worksheet `SalesData$` and columns `OrderID`, `ProductID`, `Quantity`, `SalesAmount` (or bring your own sanitized, deterministic file); see [Data Contract](docs/DATA_CONTRACT.md). The original workbook is not distributed.
6. Create the **empty lab benchmark table** using `sql/01_Benchmark_Table.sql`. DEV restore may remove this table, so recheck it before every benchmark.
7. Run `scripts/ssis/11_SSIS_SingleScript_CSharp_v3.ps1` on **64-bit Windows PowerShell 5.1** to generate and validate the `.dtsx`. It uses installed SSIS 17.0 assemblies and requires x64 ACE during generation.
8. To measure a validated package, run `scripts/ssis/12_Run_Architecture_Benchmark.ps1 -Architecture x64 -AllowResetBenchmarkTable` or `-Architecture x86` with the **appropriate working ACE architecture**. Read the truncation warning first.

### Example commands

```powershell
# In 64-bit Windows PowerShell 5.1 on the DEV/UAT VM:
.\scripts\diagnostics\Confirm_Lab_Prerequisites.ps1

# Generate + validate package (with working x64 ACE and required SSIS assemblies):
.\scripts\ssis\11_SSIS_SingleScript_CSharp_v3.ps1

# Repeat on a later run only when intentionally replacing the package:
.\scripts\ssis\11_SSIS_SingleScript_CSharp_v3.ps1 -OverwritePackage

# Explicitly destructive ONLY to the dedicated benchmark table:
.\scripts\ssis\12_Run_Architecture_Benchmark.ps1 -Architecture x86 -Runs 3 -AllowResetBenchmarkTable
```

> Do **not** use `-RunBenchmark` with the v3 generator after migrating to x86 Access Runtime. That generator requires an x64 ACE connection during metadata generation. Use the **already generated** `.dtsx` with the architecture-specific runner instead.

## Important boundaries

- Backups on an SMB share are not an independently verified, immutable off-host DR copy. Perform restore drills, protect credentials, verify storage capacity, and implement retention and alerts.
- DEV restore **overwrites** the DEV lab database; daily 05:30 restores may remove the SSIS benchmark table.
- Power BI projects and scripts do not include credentials or the original `.pbip` created on the Windows machine. Refresh and Desktop rendering require local validation.
- x86/x64 ACE Click-to-Run components are not a supported guaranteed side-by-side setup. The UAT test **migrated** x64 Access Runtime to x86 after backup. This repository intentionally does not automate that risky architecture switch.
- No passwords, database backups, Excel datasets, real PBIX files, or VM images are included.
- Deployment settings in these scripts reflect the **original lab defaults**. Parameterize and review before using them in a different project.

## Guides

[Implementation Guide](docs/IMPLEMENTATION_GUIDE.md) · [Benchmark Details](docs/BENCHMARK.md) · [Troubleshooting](docs/TROUBLESHOOTING.md) · [Reuse Checklist](docs/REUSE_CHECKLIST.md) · [Safety](docs/SAFETY.md)

## Source provenance

The core scripts are taken from files generated and revised during the original lab. The v3 generator corresponds to the successful package-generation log, and the new `12_Run_Architecture_Benchmark.ps1` is a cleaned-up reusable form of the proven 32-bit three-run block. The reconstructed runner itself has **not** been re-executed on the original Windows VM. No public GitHub commit was made to the referenced repository.

## License

MIT — see [LICENSE](LICENSE).
