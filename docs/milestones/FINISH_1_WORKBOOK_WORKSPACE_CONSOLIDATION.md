# FINISH-1 — Workbook + Workspace Consolidation

Status: complete.

This milestone closes the five-hour XML/workbook-policy refactor sequence without changing Normalizer, S-Curve math, Fake Amount, Mapping, Payment, or Rebuild engines.

## Final architecture

```text
MSP XML / P6 XML
        -> Normalize + Validate
        -> Create Progress Engine
        -> Final Workbook Policy
        -> Progress Workbook

Edited Progress Workbook
        -> Rebuild Engine
        -> same Final Workbook Policy
        -> Updated Workbook
```

## Final workbook policy

- Public working sheets remain visible: README, main, main_monthly, Payment Input, Payment, Dashboard (when present).
- Every helper/internal sheet is normal **Hidden**, not VeryHidden. This keeps development/support inspection possible from Excel while sheet protection prevents accidental edits.
- Formula/helper cells remain protected; intended user controls and inputs remain unlocked by the shared protection policy.
- Workbook structure remains unlocked.
- Snapshot workbooks keep the incremental/automatic recalculation policy.
- Live workbooks keep manual calculation with calculate-on-save; F9/Save workflow remains available.

## Desktop workspace policy

- Home opens first and explains the user workflow.
- Create Progress contains only schedule-generation actions plus Go to Mapping.
- Mapping commands (Undo, Map, Unmap, Export Mapped Workbook) appear only in Mapping.
- Mapping and Payment are optional branches; Rebuild refreshes edited workbooks.

## Deferred

- Whether Live Rebuild should restore a `progress_table` snapshot.
- Any further calculation-engine unification between Create, Snapshot Rebuild and Live Rebuild.
