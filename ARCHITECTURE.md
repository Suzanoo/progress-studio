# Progress Studio Architecture

## Purpose

Progress Studio connects a generated Progress workbook with a BOQ workbook, keeps mapping state in memory, and exports a reconciled workbook without using GUI widgets as data storage.

## Data flow

```text
Primavera XML or generated Progress workbook
        ↓
Workbook readers and validation
        ↓
Domain records
        ↓
MappingStore / MappingSession
        ├── Activity index
        ├── BOQ index
        ├── Allocation records
        ├── Search and filters
        └── Undo state
        ↓
Session repository or WorkbookExportService
        ↓
JSON session or mapped Excel workbook
```

## Sources of truth

- The generated `main` worksheet is the workbook source of truth for Activity Amount.
- `MappingStore` is the runtime source of truth for Activities, BOQ items, selections, and allocations.
- Treeview rows are presentation only. Business logic must never read values back from the GUI.
- Session JSON stores references and allocation records, not duplicated workbook data.

## Workbook contract

A Progress workbook accepted for mapping must contain a worksheet named exactly `main` and the required generated headers. Export updates Plan Activity Amount in `main`; dependent worksheets recalculate when Microsoft Excel opens the result.

## UI contract

The desktop UI is event-driven and lightweight:

- Only visible pages are rendered.
- Native Treeview selection and scrolling are preferred.
- No row tooltips, hover rendering, animation, embedded charts, or continuous repaint loops.
- Layout preferences may store only presentation state; they must not contain mapping data.

## Export contract

```text
MappingStore
    ↓ validate and reconcile
Temporary workbook
    ↓ update main + mapping sheets
OOXML package validation
    ↓ atomic replace
Final mapped workbook
```

A failed validation must not replace an existing destination workbook.

## Schedule XML import contract

Schedule XML is normalized before the workbook pipeline. The downstream schedule,
mapping, session, and export services remain source-independent.

Every non-summary activity must resolve these fields:

- Activity Name
- Plan Start
- Plan Finish

If any required value is missing, invalid, or Plan Finish is earlier than Plan Start,
the importer reports all detected issues and stops before creating a workbook.

Activity ID and WBS are optional. Missing Activity IDs are generated deterministically
as `ACT-000001`, `ACT-000002`, and so on. Missing hierarchy is represented as a flat
activity structure. Existing Microsoft Project / P6-exported XML remains supported
through the same normalized reader.
