# Benchmark — how the timing was measured

Both architectures ran a genuine SSIS 2025 Data Flow using the same saved Excel-to-SQL package, the same 100,000-row workbook and SQL destination table. Each run cleared the dedicated benchmark table, used DTExec, and verified 100,000 inserted rows with exit code 0. The clock measured **overall process invocation**, not only the internal SSIS Data Flow transform time.

| Runtime | Run 1 (s) | Run 2 (s) | Run 3 (s) | Median (s) | Mean (s) |
|---|---:|---:|---:|---:|---:|
| x64 | 3.920 | 3.524 | 3.276 | **3.524** | **3.573** |
| x86 | 3.824 | 3.991 | 4.047 | **3.991** | **3.954** |

Elapsed-time difference using x64 as baseline: `(3.991 - 3.524) / 3.524 * 100 = 13.25%`. Thus x86 took 13.25% longer by median. Another phrasing: x64 took about 11.7% less time than x86. Median x64 throughput was ~28,377 rows/s; x86 ~25,056 rows/s. All six executions passed.

**Fairness and limitations:** The architecture tests happened sequentially after changing Access Runtime, not concurrently with two installed ACE architectures. Cached data, SQL buffers, CPU contention and different process warm-up behavior may impact timings. Three runs are useful for a functional comparison but not a universal performance claim. For research-quality results, capture OS/SQL/ACE versions for each run; randomize run order if feasible; perform warm-ups; run 10–30 repetitions; isolate Agent jobs; measure CPU/memory/disk; record confidence intervals and test more file sizes and transformations.

**Original files stayed on the DEV machine:**

- `C:\SQLLab\SSIS\Results\SSIS_x64_20261016_095650.csv`
- `C:\SQLLab\SSIS\Results\SSIS_x86_20261016_165032.csv`

The repository includes `results/example-benchmark.csv`, a transcription of the observed console results—not the original CSV files and not a fresh verification.
