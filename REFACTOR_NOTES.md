# Refactor Notes

## MS-6 completed

- Replaced the final stage-07 adapter with `OkdStep` and `OkdService`.
- Moved OKD workbook logic to `infrastructure/excel/okd_workbook.py`.
- Removed `progress_studio/legacy` and `LegacyPipelineStep`.
- Removed root scripts `01` through `07` from the release package.
- Removed all subprocess usage from the application package.
- Preserved the final workbook behavior: OKD sheets are added to the accepted distribution workbook in place.
- Preserved the separate Amount Mapping copy in the project output folder.

MS-7 remains for final regression coverage, documentation cleanup, release packaging, and user-facing validation.
