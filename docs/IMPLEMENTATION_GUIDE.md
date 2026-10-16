# Implementation Guide — from zero to completed lab

This guide explains **what we did, why we did it, and what to expect**, in straightforward English. The original project ran on Windows Server 2022 and SQL Server 2025 in two domain-joined VMs. Use other names, IPs and share paths for your own project.

## Phase 0 — Choose the lab boundaries

Create a **PROD simulation** and a separate **DEV/UAT simulation**. They need not contain real production data. In our setup, the original hosts were `DCSQLPRI01` (192.168.88.18) and `DCSQLDB-SEC1` (192.168.88.19). The database names were `PowerBIAnalytics_PROD` and `PowerBIAnalytics_DEV`. Domain Windows authentication was used. Place each VM on reliable storage and ensure you can restore it if a test fails.

We separate these roles so backups are taken from a stable source and restoration can be tested without damaging the source.

## Phase 1 — Create the reporting dataset on PROD

The PowerShell script `scripts/prod/01_PROD_Setup_Backup.ps1` creates a marked synthetic dataset with Customers, Products, SalesOrders, SalesOrderDetails, Payments, and a reporting view `dbo.vw_SalesAnalytics`. It uses fixed lab database names on purpose to reduce the chance of damaging real data.

On the intended PROD **lab** server, read the script and run:

```powershell
.\01_PROD_Setup_Backup.ps1 -RunBackupNow
```

Inspect the row counts (original test: 1,000 customers, 100 products, 20,000 orders and 19,000 payments). The view supplies sales amounts, dates, products and regions for business intelligence.

**Why:** A predictable dataset makes it possible to test recovery and reporting repeatedly without exposing customer data.

## Phase 2 — Back up PROD and refresh DEV

The PROD script creates a local backup folder (default `C:\DBBACKUP\Full`) and configures a SQL Agent job for 05:00. Its backup command uses compression, checksums, and `RESTORE VERIFYONLY`. Access the folder through an **authorized** SMB share (original example `\\192.168.88.18\DBBACKUP$`). Restrict both SMB and NTFS permissions; a share name ending in `$` merely hides it from casual browsing and is not a security boundary.

The DEV script `scripts/dev/02_DEV_Restore.ps1` selects a backup from the current date, verifies the backup, and restores into **PowerBIAnalytics_DEV**, not PROD. It also supports a scheduled 05:30 SQL Agent job.

```powershell
# Save on DEV at C:\SQLLab\02_DEV_Restore.ps1 before job installation.
.\02_DEV_Restore.ps1 -RunOnce -ApproveOverwrite
# Inspect the source and then install job:
.\02_DEV_Restore.ps1 -InstallJob -ApproveOverwrite
```

**Why:** Developers should work with a repeatable copy, and backup success alone is insufficient without a restore test. The original manual PROD/DEV jobs succeeded; the scheduled jobs were enabled. Ongoing daily reliability, alerting, retention and cross-host disaster recovery were *not* proven.

**Risk:** Restores overwrite the DEV database. They may remove the benchmark table in Phase 4. Do not reuse the script without adapting and testing its safety restrictions.

## Phase 3 — Build a Power BI dashboard

Run `scripts/powerbi/07_EndToEnd_PowerBI.ps1` where Power BI Desktop is installed, using SQL read permissions:

```powershell
.\07_EndToEnd_PowerBI.ps1 -SqlServer YOUR_PROD_LAB_HOST -OutputRoot C:\SQLLab\PBI_Automated
```

Open the generated `.pbip`, provide your own SQL credentials if prompted, refresh, and inspect the seven visuals. The earlier lab showed about 11.05 million in synthetic revenue, 20,000 orders and 1,000 customers. **Generated visuals and reported figures are not substitutes for a refreshed data-model check**. The project is for Power BI Desktop; gateway and cloud publication are separate tasks.

**Why:** It shows the journey from SQL data to a human-readable business dashboard.

## Phase 4 — Create a real SSIS pipeline

Install both SSIS 17 runtimes, and use **64-bit Windows PowerShell 5.1** to run `scripts/ssis/11_SSIS_SingleScript_CSharp_v3.ps1`. The original setup also had Visual Studio 2026 SSIS 170 assemblies; the generator searches version-specific component ProgIDs. It uses a typed C# helper inside PowerShell, avoiding errors from direct dynamic COM access.

Prepare an actual Excel workbook at `C:\SQLLab\SSIS\Input\SalesBenchmark_100K.xlsx`, with worksheet `SalesData$` and the four columns from [Data Contract](DATA_CONTRACT.md). The actual workbook is **not bundled**. Also create the destination table using `sql/01_Benchmark_Table.sql` on the DEV lab database.

```powershell
.\11_SSIS_SingleScript_CSharp_v3.ps1
# It should emit a DTSX and run DTExec /Validate, exit 0, no data load.
```

Do not treat successful validation as data-load success. In the actual lab, version-specific ProgIDs `DTSAdapter.ExcelSource.9` and `DTSAdapter.OleDbDestination.9` solved a component initialization problem. Another SSIS installation may use other registered values.

**Why:** A genuine SSIS Data Flow measures Excel Source and SQL Destination execution, unlike a PowerShell CSV import pretending to be SSIS.

## Phase 5 — Execute the x64 and x86 benchmark

You need working, **architecture-matched** ACE 16 and MSOLEDBSQL providers. The original x64 test ran first. The team then changed Microsoft 365 Access Runtime from x64 to x86 using the Microsoft Office Deployment Tool and a VM backup. This repository does not automate that switch because it can interrupt applications.

Use the same saved DTSX and Excel input on both tests:

```powershell
# Deliberately clears only the dedicated benchmark table.
.\12_Run_Architecture_Benchmark.ps1 -Architecture x64 -Runs 3 -AllowResetBenchmarkTable
.\12_Run_Architecture_Benchmark.ps1 -Architecture x86 -Runs 3 -AllowResetBenchmarkTable
```

For each run, the test truncates the benchmark table, starts a stopwatch, executes the package, records the native process exit code, counts SQL rows, and saves a CSV/log. A run passes only when exit code is zero **and** the row count matches the expected 100,000. Do not execute during the 05:30 refresh window.

**Why:** Identical workloads and validated rows give us a better comparison than relying on execution time alone.

## Phase 6 — What makes it reusable

For another project, replace the **synthetic dataset generator, column schema, workbook and DTSX mappings**, customize the SQL hosts/share, choose the job schedule, add IAM least privilege, and design retention/DR. Read [Reuse Checklist](REUSE_CHECKLIST.md) before making changes.

## What is intentionally unfinished

A separate immutable/off-site backup, audited restore drills, production monitoring, alerting, SQL credentials secret management, Power BI Service deployment, CI tests on Windows, and higher-volume statistically meaningful benchmarks. These are important next steps, not delivered features.
