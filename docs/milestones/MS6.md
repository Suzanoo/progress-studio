# Progress Studio V3 — MS-6 Final Workbook Export

## Goal

Create a safe final mapped Progress workbook from the in-memory Mapping Store. Export never reads values back from Tkinter tables.

## Completed

- Added a dedicated `WorkbookExportService` and `MappedWorkbookExporter`.
- Added pre-export reconciliation for Activities, BOQ items, allocations, mapped money, remaining money, and item status.
- Incomplete mappings require explicit confirmation before export.
- The source Progress workbook is never overwritten.
- Existing output files require an explicit overwrite decision from the desktop save dialog.
- Export is written to a temporary workbook in the output folder and atomically replaced only after a successful save.
- Existing worksheets and formulas are preserved.
- `Amount Mapping` Activity amounts and statuses are updated from allocation records.
- `BOQ Activity Mapping` is rebuilt as a formatted Excel table.
- `Mapping Summary` is added as the first worksheet with reconciliation totals.
- Excel full recalculation is requested when the exported workbook opens.
- A stable `BOQ ID` is generated for export traceability while the existing BOQ key remains available for session compatibility.

## Final workbook sheets

### Mapping Summary

Contains:

- Complete or Partial export status
- Activity and BOQ counts
- Allocation record count
- Full, Partial, and Unmapped BOQ item counts
- Total BOQ amount
- Allocated amount
- Remaining amount
- Allocated percentage

### Amount Mapping

For each Activity row:

- `Amount` is replaced with the allocated BOQ amount.
- `Status` is set to `MAPPED` or `UNMAPPED` when that column exists.
- WBS rows and unrelated workbook structure remain unchanged.

### BOQ Activity Mapping

Columns:

```text
Activity ID
BOQ Key
Source Sheet
Source Row
WBS-2
WBS-3
WBS-4
BOQ Description
BOQ Amount
Share %
Allocated Amount
Mapping ID
BOQ ID
```

The first eleven columns preserve the MS-4 export contract. `Mapping ID` and stable `BOQ ID` are appended for traceability.

## Output policy

Default name:

```text
<progress-workbook-name>_mapped.xlsx
```

Rules:

- Output must use `.xlsx`.
- Output cannot be the loaded source workbook.
- The GUI asks before replacing an existing file.
- A failed export does not leave a partial final workbook.

## Partial mapping

Partial export is allowed only after explicit user confirmation. The exported `Mapping Summary` clearly records `Export status = Partial` and the remaining reconciliation values.

## Run

```powershell
python desktop.py
```

## Tests

```powershell
pytest -q
```
