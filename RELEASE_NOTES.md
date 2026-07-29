# Release Notes

## 2.0.4

- Add an embedded desktop S-curve preview sourced from the generated progress workbook.
- Refresh the chart after every successful run so distribution and cutoff changes are visible immediately.
- Keep chart rendering isolated from workbook generation; a chart error does not invalidate output.

## 2.0.1

- Removed obsolete root-level `excel_toolkit/`.
- Moved Activity ID rules into the domain layer.
- Moved distribution algorithms and auto-selection rules into the services layer.
- Moved Excel theme helpers into Excel infrastructure.
- Preserved the approved end-to-end regression manifest.

# Progress Studio 2.0.0

## Release status

Progress Studio V2 is the completed architectural refactor of V1.

## Completed changes

- Replaced procedural scripts `01` through `07` with seven in-process pipeline steps.
- Added application, domain, service, infrastructure, pipeline, presentation, and configuration layers.
- Renamed the primary imported schedule worksheet to `main`.
- Removed subprocess-based stage execution and legacy adapters.
- Centralized settings and workbook schema.
- Kept domain models independent from `openpyxl`.
- Added automated architecture, CLI, error-handling, and end-to-end regression tests.
- Updated all user-facing application text and release documentation to English.

## Approved example regression manifest

- WBS nodes: 82
- Activities: 172
- Weekly columns: 76
- OKD data rows: 510
- OKD formula links checked: 40,290
- Final worksheet row counts:
  - `main`: 519
  - `Distribution Report`: 179
  - `progress`: 77
  - `progress_table`: 511 including the header

## Compatibility note

The workbook contains live Excel formulas. Open and save the generated workbook in Microsoft Excel before uploading it to OKD.

## 2.0.3

- Generate sequential child WBS codes for Activity rows in `main`.
- Reuse the same `ActivityWbsSequencer` rule used by `progress_table`.
- Example: parent `1.2` produces `1.2.1`, `1.2.2`, and so on.
- Preserve date formatting for `progress_table` weekly headers.

## 2.0.2

- Fixed missing WBS values on Activity rows in the `main` worksheet.
- Fixed `progress_table` weekly headers displaying Excel serial numbers instead of dates.
- Added regression tests for both fixes.

## V3 MS-2

- Added cascading WBS-2/WBS-3 BOQ filters.
- Kept paginated, memory-backed mapping tables.
- Added golden Progress and BOQ workbooks for regression testing.

## V3 MS-3

- Added the full-allocation BOQ-to-Activity mapping engine.
- Added BOQ allocated amount, remaining amount, mapping status, and mapped Activity display.
- Added exact single-command undo and unmap behavior.
- Updated only affected visible rows after mapping commands.
- Removed the unsafe Clear All action pending session recovery support.

## V3 MS-4

- Added percentage-based BOQ allocation across multiple Activities.
- Added over-allocation validation so combined BOQ shares cannot exceed 100%.
- Added live Partial/Full status and remaining BOQ amount calculations.
- Added pair-specific unmap and exact share undo.
- Added Share % and Allocated Amount columns to mapping export.

### MS-4 UI refinement
- Changed the BOQ table `Remaining` display from money to remaining allocation percentage.
- Kept `Amount` and `Allocated` as monetary values.
- Added store-level percentage calculation and tests.
