# Desktop S-Curve Acceptance

- The stable Phase 2 implementation remains available on branch `feat/desktop-phase2` at commit `af25258`.
- Desktop runs on branch `feat/desktop-scurve`.
- After each successful pipeline run, the app reads activity rows from the generated `main` sheet and refreshes the chart.
- Plan is shown as cumulative weighted progress using Activity Amount.
- Actual is shown only through the latest week containing actual input.
- Running again with a different cutoff day or distribution replaces the prior preview.
- Chart failure does not invalidate the generated workbook or disable the open-file buttons.
