# MS-P1.9 — Dashboard UI and Reporting Behavior

## Goal

Keep the Excel Dashboard simple and make reporting behavior match planning practice.

## Dashboard controls

- `View`: Weekly / Monthly
- `Cutoff Date`: weekly reporting date
- KPI calculations always use the selected cutoff date.
- Weekly/Monthly changes the S-Curve display period only.

## KPI cards

Exactly four KPI cards are shown:

1. Planned Progress
2. Actual Progress
3. Schedule Status
4. Time Impact

Progress Gap and Reporting Period cards were removed.

## S-Curve contract

- Plan curve always displays the full baseline through project finish.
- Actual curve is displayed only through the selected cutoff date.
- Planned KPI is calculated at the cutoff date; it does not use the final baseline value.
- Actual KPI is calculated at the cutoff date.

## Theme configuration

Dashboard visual settings are in:

`progress_studio/config/dashboard_theme.json`

Configurable values include colors, chart grid/axis colors, chart size/style, line widths, markers, and activity row count.
