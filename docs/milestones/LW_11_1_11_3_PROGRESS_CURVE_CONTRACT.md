# LW-11.1 → LW-11.3 — Progress Curve Contract

## Goal

Make `main` the editable source of truth, `progress` the sole S-Curve calculation layer, and `Dashboard_Data` a presentation-only Weekly/Monthly adapter.

## LW-11.1 — Progress Contract

- `main` S-Curve rows no longer reference Dashboard cutoff controls.
- `progress` is rebuilt as a tiny hidden worksheet with only `Date`, `Plan`, and `Actual`.
- Plan always exposes the full cumulative baseline.
- Actual is blank when `Date > Dashboard!K5`.
- Excel Manual calculation + calculate-on-save remains the Live Workbook policy; F9/Save refreshes the contract.

## LW-11.2 — Monthly Curve Adapter

- Dashboard monthly chart data no longer reads `main_monthly` S-Curve rows.
- Monthly chart points are the last weekly `progress` point in each calendar month.
- `main_monthly` remains an Activity Data monthly view and has no Dashboard cutoff ownership.

## LW-11.3 — Dashboard Dumb Renderer

- `Dashboard_Data` reads Weekly Plan/Actual from `progress`.
- Monthly Plan/Actual are references to month-end `progress` rows.
- Selected Plan/Actual formulas only switch Weekly/Monthly view.
- No second Actual cutoff condition exists in `Dashboard_Data`.
- Activity Table remains direct-to-main until LW-11.4.

## Recalculation Flow

`main` → F9/Save → `progress` → `Dashboard_Data` → Dashboard chart/KPI.

Structural edits that change workbook shape remain Rebuild responsibilities.
