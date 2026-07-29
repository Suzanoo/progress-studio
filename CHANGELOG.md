# Changelog

## MS-6.1 — Excel export compatibility hotfix

- Fixed Excel repair warning caused by duplicate worksheet and Table AutoFilters.
- Added direct OOXML table and relationship validation before atomic export replacement.
- Added empty-allocation and Excel compatibility regression tests.
- Consolidated milestone and acceptance documentation under `docs/`.
- Removed obsolete phase acceptance notes and premature MS-7/MS-8 acceptance files.

# Release Notes

## V3 MS-6 — Final Workbook Export

- Added validated, atomic mapped-workbook export.
- Added Mapping Summary reconciliation worksheet.
- Added partial-export confirmation and final status reporting.
- Preserved formulas and existing workbook sheets.
- Added formatted BOQ Activity Mapping table with Mapping ID and stable BOQ ID.
- Prevented accidental source-workbook overwrite.
- Added full MS-6 export tests.

# MS-5 Session Hardening

- Added session schema version 2 and v1-to-v2 migration.
- Stored workbook filename alongside saved absolute path and SHA-256.
- Added safe relink flow for moved or renamed workbooks.
- Rejected relink candidates whose content differs from the saved fingerprint.
- Added clearer validation messages and migration/relink tests.
- Documented that Undo history starts fresh after a session is loaded.

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

## V3 MS-5 — Persistent Mapping Session

### Added

- Atomic JSON mapping-session persistence.
- Save Session, Load Session, and Recent Sessions desktop actions.
- Auto-save after mapping changes once a session has been saved.
- SHA-256 validation of Progress and BOQ workbooks during restore.
- Allocation restoration validation.
- Confirmed Clear All command with one-step Undo recovery.
- `README_ROADMAP.md` and `COPILOT.md` for future developers and coding agents.

## [2.1.3] - 2026-07-29

### Fixed

- Validate the required `main` worksheet and generated headers when loading Progress workbooks.
- Update mapped Activity Amount values in `main` instead of treating `Amount Mapping` as the export source of truth.
- Rebuild WBS and Project Summary Amount formulas in `main` and keep Actual-row Amount blank.
- Reject exports whose Activity Amount total does not reconcile with the allocated BOQ amount.
- Keep the Activity and BOQ table vertical scrollbars visible during resize.

### Changed

- Deferred S-Curve work outside the current V3 roadmap.
- Replaced the former MS-7 S-Curve milestone with Mapping Workspace UX.
- Added a user-manual TODO describing the Progress workbook contract.

## [2.2.0] - 2026-07-29

### Added

- Added a Focus Mapping workspace mode that collapses the Primavera generator panel.
- Added collapsible Workbook Inputs with automatic collapse after loading.
- Added optional persisted layout preferences for collapsed states and the mapping divider.
- Added `ARCHITECTURE.md` and MS-7 milestone/acceptance documentation.

### Changed

- Start the desktop window maximized with a platform-safe fallback.
- Compact mapping controls and allocate more table width to descriptions.

### Performance

- Added no hover tracking, row tooltips, animation, charts, or continuous repaint loops.

## [2.1.4] - 2026-07-29

### Fixed

- Force Microsoft Excel to rebuild formulas after mapped Activity Amount values are written to `main`.
- Reset the calculation engine ID and remove reliance on stale formula caches.
- Add OOXML regression checks for calculation properties and calculation-chain removal.

### Removed

- Remove the embedded S-Curve preview and its unused desktop service from the current V3 scope.

### Documented

- Clarify that Activity weekly percentages are manual distribution inputs, while WBS and Project Summary percentages are amount-weighted formulas.
