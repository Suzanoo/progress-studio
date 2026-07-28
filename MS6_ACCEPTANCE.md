# MS-6 Acceptance — OKD and Legacy Removal

MS-6 passes only when all conditions below are true:

1. `OkdStep` and `OkdService` are part of the application pipeline.
2. OKD Excel logic is under `infrastructure/excel/okd_workbook.py`.
3. `progress` and `progress_table` are generated without subprocesses.
4. The temporary legacy adapter and `progress_studio/legacy` package are removed.
5. Root scripts `01` through `07` are removed from the release package.
6. No application module imports or calls `subprocess`.
7. The final workbook remains the distribution workbook with OKD sheets added in place.
8. The Amount Mapping workbook is copied to the project output folder.
9. No Thai text exists in application source.
10. Automated MS-6 tests pass.
