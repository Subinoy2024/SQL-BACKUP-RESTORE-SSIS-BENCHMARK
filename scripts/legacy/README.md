# Historical / superseded experiments

These scripts were generated during earlier investigation and some were replaced after runtime failures. They are included for educational traceability and comparison only. **Do not treat them as validated deployment entrypoints.** Use `scripts/prod`, `scripts/dev`, `scripts/powerbi`, and `scripts/ssis` first, and review all scripts before running.

The initial SSIS generator failed in PowerShell COM calls. The early typed C# generator required multiple fixes; the v3 generator in the main `scripts/ssis` folder is the version which generated, validated and ran successfully in the user's UAT setup.
