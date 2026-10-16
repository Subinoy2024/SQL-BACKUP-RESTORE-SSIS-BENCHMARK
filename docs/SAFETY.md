# Safety, access and rollback

- **Simulation first:** PROD/DEV labels refer to lab roles. Never point the scripts at an existing business database without a separate reviewed change plan.
- **Restore is destructive:** `02_DEV_Restore.ps1 -ApproveOverwrite` replaces `PowerBIAnalytics_DEV`. Confirm backup metadata and recovery paths.
- **Benchmark is destructive to one table:** `12_Run_Architecture_Benchmark.ps1 -AllowResetBenchmarkTable` truncates only `dbo.SSIS_Benchmark_Results`, but ensure the saved DTSX actually writes to the intended database and table.
- **Scheduled job overlap:** 05:00 backup and 05:30 restore are examples. Confirm timezone, run durations, DST policy, and backup freshness. The benchmark script blocks 04:30–06:30 by default.
- **Backups:** `RESTORE VERIFYONLY` helps check readability but is not a complete restore drill or an off-host/immutable copy. Use separate storage, retention, alerts and periodic test restores.
- **Permissions:** least-privilege Windows/SQL accounts; read-only reporting permissions; restricted SMB and NTFS ACLs. Hidden shares (`$`) are not secure by themselves.
- **Office architecture:** Do not install x86 Access Runtime alongside x64 Office by force. The original lab migrated one Click-to-Run architecture at a time after a Proxmox backup. This can break x64 ACE consumers. No migration automation is shipped.
- **Security:** No passwords, connection tokens, customer data or certificate private keys should be committed. Windows Integrated authentication is used in the examples.
- **Environment fit:** Scripts intentionally reference SQL 2025, Visual Studio 2026, Windows PowerShell 5.1 and original lab paths. Test your adapted version before production use.
