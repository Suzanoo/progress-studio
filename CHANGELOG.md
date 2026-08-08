## MS-R1 — Self-contained Rebuild Workbook

- Upgraded `.progressstudio` project schema from v7 to v8.
- Saved projects now embed verified Progress and BOQ workbook snapshots.
- A v8 project can restore its sources from a local integrity-checked cache when the original files are missing, renamed, or on another machine.
- Export workspace now presents **Rebuild Latest Workbook**, reusing the current generation/export engine and saved mapping/tree state.
- Legacy v7 and older projects migrate safely but require one successful source relink/open and Save before standalone rebuild is available.
- Actual Progress migration from an edited legacy workbook remains explicitly out of scope for MS-R1.

## MS-P1.8 — Dashboard Data and Theme Refactor


## MS-P1.14 — Progress-Source Dashboard

- Dashboard KPI Plan/Actual now read the stable `progress` sheet directly at the selected cutoff.
- Time Impact = rounded baseline project duration (`project_finish - project_start`) × absolute schedule variance.
- Schedule Status and Time Impact cards use delay/ahead/on-schedule conditional colors.
- `Dashboard_Data` remains a thin Weekly/Monthly chart adapter instead of duplicating schedule business logic.
- Activity Progress now mirrors the main-sheet native Excel outline hierarchy for Project → WBS → Activity drill-down.

## MS-P1.13 — Dashboard + Main Filter Polish

- Dashboard Activity Progress now renders every progress row instead of stopping at eight.
- Removed the Status column; the exception table now exposes only the WBS filter.
- Main worksheet keeps a full AutoFilter range but shows filter buttons only on Row Type and P/A.

- Fixed blank `Dashboard_Data` when `progress.week_start` contains worksheet-reference formulas.
- Added dynamic progress-header discovery and date parsing.
- Added a clear generation error when weekly dashboard data cannot be built.
- Cleaned up Dashboard controls and initialized the cutoff date with a real value.
- Added editable Dashboard theme config at `progress_studio/config/dashboard_theme.json`.
- Dashboard remains generated before BOQ mapping.

# Changelog

## MS-P1.6 — Excel Dashboard

- Added a separate `Dashboard` worksheet as the first workbook tab.
- Added Weekly / Monthly and Cutoff Date dropdown controls.
- Added formula-linked KPIs, S-Curve, and Activity Progress summary.
- Added hidden `Dashboard_Data` helper sheet while preserving `main`, `progress`, and `progress_table`.
- Integrated dashboard generation into both fresh workbook generation and mapped workbook export.
- Added automated dashboard and generation-progress tests.

## MS-P1.2 — Cached Workbook Identity

- Cache Progress and BOQ workbook identities at load/relink boundaries.
- Autosave now writes only the project JSON and no longer reopens or hashes Excel workbooks on every mapping action.
- Manual Save and Save As reuse the verified in-memory identities.
- Preserve safe relink and workbook mismatch protection from MS-P1.1.
- Add regression coverage proving cached session creation performs no workbook fingerprinting.

## MS-P1.1 — Workbook Identity & Safe Relink

- Replaced binary-only workbook matching for new projects with a stable Excel semantic identity.
- Allowed moved, renamed, re-saved, and formatting-only workbook changes to relink safely.
- Continued rejecting worksheet data or formula changes that may invalidate mappings.
- Preserved strict SHA-256 validation for legacy project sessions until they are saved again.
- Upgraded the mapping session schema from version 6 to version 7 with automatic migration.
- Updated relink guidance and added regression tests for Excel re-save, formatting, rename, mismatch, and legacy sessions.

## MS9.1 — Production desktop UI foundation

- Rebuilt the desktop shell around the approved production mockup: menu bar, navigation sidebar, six-stage workflow header, mapping-first workspace, and application status bar.
- Preserved the existing mapping, session, allocation, export, and workbook-generation behavior.
- Centralized the visual design system in `presentation/gui/theme.py`.
- Centralized new English-first interface text in `presentation/gui/strings.py` for future localization.
- Fixed release-version drift between desktop settings and `progress_studio.version`.
- Added MS9.1 architecture tests and a headless desktop launch smoke test.

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

## [2.3.0] - 2026-07-29

### Added

- Source-independent Schedule XML reader with namespace-safe field resolution.
- Strict Activity Name, Plan Start, and Plan Finish import contract.
- Aggregated validation errors that stop processing before workbook creation.
- Deterministic generated Activity IDs when source IDs are unavailable.
- Generic, namespace, Unicode, missing-field, malformed XML, and regression fixtures.

### Changed

- Renamed the desktop input from Primavera XML to Schedule XML.
- Kept the former `PrimaveraXmlReader` as a compatibility adapter.
- Isolated XML parsing and validation from the existing workbook, mapping, session, and export pipeline.

## MS-P1.3 Combined Theme Refactor

- Restored WBS level color hierarchy in the main-sheet timescale from the original MS-P1.3 build.
- Kept the separate WBS level hierarchy in the Activity Data section.
- Removed the dark top border from WBS level 1 Activity Data rows.
- Centralized export color configuration in `progress_studio/infrastructure/excel/export_theme.py`.

## MS-P1.5 - Actual Amount formulas

- Calculate Activity Actual Amount from Plan Amount and Actual `% Complete`.
- Roll up Actual Amount to WBS and Project Summary rows.
- Keep weekly Actual progress weighted by full Plan Amount.
- Show Actual Amount cells with the normal currency format instead of hiding them.

## MS-P1.10 - Embedded Dashboard KPI Icons

- Add four lightweight transparent PNG icons for Planned Progress, Actual Progress, Schedule Status, and Time Impact.
- Embed KPI icons inside the generated XLSX so they remain visible on other computers without extra fonts or internet access.
- Add dashboard icon settings to `progress_studio/config/dashboard_theme.json` for enable/disable, size, and asset filenames.
- Add regression coverage that verifies all four icons survive workbook save/reopen.

## MS-P1.18 - Workbook Performance

### Changed

- Use Excel automatic dependency calculation instead of forcing a full workbook recalculation on every open.
- Keep `fullCalcOnLoad` and `forceFullCalc` disabled for normal generation/export while retaining an explicit full-rebuild escape hatch for repair/debug workflows.
- Store generated Activity Plan weekly values as static values in `progress_table`; Actual rows, Amount links, and WBS/Project rollups remain live.
- Store generated Activity Plan monthly values as static values in `main_monthly`; Actual and summary rows remain formula-driven.
- Preserve a non-zero Excel calculation engine ID and request normal recalculation on save.
