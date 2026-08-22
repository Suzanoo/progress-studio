# MS-2 Acceptance — Import and Schedule

## Scope

MS-2 replaces legacy stages 01 and 02 with application pipeline components.
No feature, formula, workbook layout, or business-rule change is included.

## Pass criteria

- [x] `Activity` and `ScheduleWindow` domain models exist.
- [x] Primavera XML parsing is implemented by `PrimaveraXmlReader`.
- [x] WBS date roll-up is implemented by `ScheduleService`.
- [x] Workbook export is isolated under `infrastructure/excel`.
- [x] `ImportStep` and `ScheduleStep` are part of the application pipeline.
- [x] The legacy continuation begins at stage 03.
- [x] Stages 01 and 02 no longer run through subprocess calls.
- [x] V2 stage-01 workbook matches the V1 baseline cell-by-cell.
- [x] V2 stage-02 workbook matches the V1 baseline cell-by-cell.
- [x] All MS-1 and MS-2 automated tests pass.
- [x] Source code and documentation contain no Thai text.

## Test command

```powershell
python -m unittest discover -s tests -v
```

Expected result:

```text
Ran 7 tests
OK
```

## Remaining legacy boundary

Scripts `03_build_timescale.py` through `07_build_okd_sheets.py` remain behind the temporary legacy continuation. They are replaced in MS-3 through MS-6.
