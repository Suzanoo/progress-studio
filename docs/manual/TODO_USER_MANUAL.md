# TODO — User Manual

This file collects user-facing behavior that must be explained before the production manual is released.

## Progress workbook contract

- The Progress workbook is generated with a worksheet named exactly `main`.
- Do not rename or delete the `main` worksheet.
- `main` is the source of truth for Activity Amount during mapped-workbook export.
- The exporter updates Plan Activity Amount values in `main` by matching `Activity ID`.
- WBS and Project Summary Amount cells in `main` are rebuilt as formulas from descendant Plan Activity rows.
- Actual-row Amount cells remain blank.
- `progress` and `progress_table` are downstream worksheets and are expected to receive updated values through formulas linked to `main` when Excel recalculates.
- If `main` is missing, Progress Studio stops loading and asks the user to rename the worksheet back to `main`.
- If required headers in `main` were renamed or deleted, the user must regenerate the Progress workbook or restore the original headers.

Required `main` headers:

```text
Row Type
Activity ID
P/A
Outline Level
Amount
```

## Export behavior

- Export never writes Amount directly into `progress` or `progress_table`.
- Excel is requested to perform a full recalculation when the exported workbook is opened.
- The sum of Activity Amount written into `main` must reconcile with the allocated BOQ amount; otherwise export fails without replacing the destination file.

## Deferred topics

- S-Curve generation and preview are outside the current V3 roadmap and must not be presented as an available MS-7 feature.
- Workbook relocation and session relinking are documented under the session workflow.
