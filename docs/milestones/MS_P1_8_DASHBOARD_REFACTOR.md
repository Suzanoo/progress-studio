# MS-P1.8 — Dashboard Data and Theme Refactor

## Purpose

Create a usable Excel Dashboard during the progress-generation stage, before BOQ mapping.

## Data contract

- `progress` is the weekly source of truth.
- `Dashboard_Data` resolves weekly dates from direct values, ISO date strings, or simple worksheet-reference formulas.
- Weekly Plan and Actual remain live links to `progress` so Excel recalculates after mapping.
- Monthly values are derived from the final weekly value in each month.
- Generation fails clearly when no valid weekly dates are found.

## Dashboard controls

- View: Weekly / Monthly
- Cutoff Date: weekly date list
- Default cutoff: latest available weekly date

## Theme configuration

Dashboard colors, font, title, chart size, and activity row count are stored in:

`progress_studio/config/dashboard_theme.json`

Hex colors do not use `#` because they are passed directly to openpyxl.
