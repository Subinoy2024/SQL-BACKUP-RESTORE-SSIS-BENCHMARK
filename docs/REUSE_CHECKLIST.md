# Reuse this project for another customer or lab

1. Start with a fresh repo/fork and choose non-production hosts.
2. Inventory SQL edition/version, Windows identity permissions, Agent availability, SSIS runtimes, ACE and OLE DB drivers.
3. Replace hard-coded names, IP addresses and shares with your approved configuration. Some scripts deliberately reject other database names; change those protections carefully, not by accident.
4. Review schema and sample-data volume; keep production data out of source control.
5. Decide backup destination, ACLs, encrypted storage, retention, verification and recovery targets (RPO/RTO).
6. Validate a one-time backup and restore before enabling any schedules.
7. Build the Power BI report locally; verify desktop refresh, model measures, visual values and row-level permissions.
8. Prepare a deterministic Excel workbook and destination table with schema matching the generator.
9. Generate SSIS package with an appropriate x64 environment; inspect connection managers and map columns.
10. Run each architecture only with a matching ACE provider, using a reversible change plan.
11. Use separate CSV/logs per run and verify exit codes and row counts; never invent missing benchmark results.
12. Add CI validation, secret scanning, Windows integration tests, monitoring and recovery rehearsals before broader adoption.
