# Progress Studio V3 — MS-2 High Performance Mapping UI

## Completed

- Independent Progress and BOQ workbook loading retained from MS-1.
- Paginated Activity and BOQ tables.
- WBS hierarchy headers on the Activity side.
- Lightweight checkbox glyphs (`☐` / `☑`) instead of per-row widgets.
- Search runs only on Search/Enter.
- Cascading BOQ filters for WBS-2 and WBS-3.
- Indexed WBS filtering in `MappingStore`.
- User-selected BOQ worksheet; no auto-detection.
- Golden test workbooks included under `example/golden/`.

## Not included yet

- Share allocation.
- BOQ remaining balance.
- Save/load mapping session.
- Final mapped workbook workflow changes.
- S-Curve redesign.

## Run

```powershell
python desktop.py
```
