# MS-P1.6 — Excel Dashboard

## Goal

Add a simple vertical Excel dashboard to every generated Progress Studio workbook without changing the existing source-sheet contracts.

## Workbook structure

- `Dashboard` — visible first worksheet for reporting and review.
- `Dashboard_Data` — hidden helper worksheet for Weekly/Monthly chart switching.
- `main`, `progress`, and `progress_table` — remain the live source worksheets.

## Dashboard controls

- **View** dropdown: Weekly / Monthly.
- **Cutoff Date** dropdown: available weekly reporting dates.

## Dashboard content

- Planned Progress
- Actual Progress
- Schedule Status
- Time Impact
- Progress Gap
- Reporting Period
- S-Curve (Plan vs Actual)
- Activity Progress summary

## Architecture

Excel-specific layout and formulas are isolated in:

- `progress_studio/infrastructure/excel/dashboard_workbook.py`
- `progress_studio/services/dashboard_service.py`

`WorkbookGenerationService` builds the dashboard after OKD sheets. `MappedWorkbookExporter` refreshes it again after the final amount and mapping sheets are written.
