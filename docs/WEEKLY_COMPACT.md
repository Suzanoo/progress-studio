# Weekly Compact — accepted presentation contract

## Authority and acceptance

- Weekly Compact V1: `main @ c9087ea`, tag `weekly-compact-v1-stable`.
- Final Presentation: `main @ 17cc3f3`, tag `weekly-compact-final-presentation-stable`.
- Product Owner Desktop Excel acceptance: PASS (Open + F9; Actual edit + F9;
  cutoff changes and collision; WBS fold/unfold; Save/Close/Reopen).
- Windows full regression: 839 passed, 6 skipped, 6 subtests passed, 0 failed.

## Source and presentation ownership

| Element | Authoritative source |
|---|---|
| Physical X geometry | Weekly `main` reporting columns |
| Gantt body / P-A rows | `main_monthly` |
| Bottom monthly data (P/AP/A/AA) | `main_monthly` |
| S-Curve | All retained Weekly Plan/Actual points |
| Monthly checkpoints | Last retained Weekly reporting point in each month |
| Current-cutoff Plan/Actual markers | Weekly values at selected Weekly cutoff |
| Red cutoff line | Shared Weekly cutoff |
| Editable cutoff | `main` (`PS_WEEKLY_OVERLAY_CUTOFF`) |
| Compact cutoff date | Locked linked display after footer |

Monthly body/footer remain monthly totals and accumulations even if the Weekly
cutoff occurs earlier in that month. They must not be silently converted to
"as of cutoff" figures. When a monthly checkpoint coincides with the selected
cutoff, the purple current-cutoff marker/label takes precedence; suppress the
monthly duplicate. Preserve the full Weekly S-Curve, not a resampled Monthly
curve. There is no independent Compact cutoff or new Monthly calculation model.

## Implementation boundary

- `weekly_compact_workbook.py`: Compact presentation coordination.
- `weekly_compact_monthly_workbook.py`: semantic projection of monthly body/footer.
- `traditional_overlay_workbook.py`: sparse chart helper series and markers.
- Excel adapter reserves `Dashboard_Data` AC:AG for Compact chart helpers.
- Generated source references are resolved from reporting structure; fixed
  helper columns belong to the Excel adapter, not domain/services.
- Do not introduce VBA, a separate date-axis/scatter engine, or raw OOXML
  manipulation merely to implement this accepted presentation.

## Desktop calculation and acceptance

Excel may require F9 after initial open before the generated monthly cells and
S-Curve render. The accepted Desktop test explicitly checked Open + F9 rather
than classifying the pre-calculation blank display as a failure. Actual edits
recalculate via F9; Save/Close/Reopen must preserve presentation and drawings
without repair warnings. Changes requiring Python-owned regeneration still
require their owning Rebuild path.

This is source/workbook acceptance, not a claim of a newly distributed build.
