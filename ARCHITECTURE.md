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
Self-contained `.progressstudio` project or rebuilt Excel workbook
```

## Sources of truth

- The generated `main` worksheet is the workbook source of truth for Activity Amount.
- `main` is also the editable weekly timescale source; `main_monthly` is a formula-derived monthly presentation view and never owns progress inputs.
- `MappingStore` is the runtime source of truth for Activities, BOQ items, selections, and allocations.
- Treeview rows are presentation only. Business logic must never read values back from the GUI.
- `.progressstudio` v8 stores mapping/tree state plus verified embedded copies of the Progress and BOQ source workbooks. This makes a saved project self-contained for future workbook rebuilds.
- Workbook snapshots are source preservation only; runtime business logic still reads normalized domain records and never reads GUI widgets.

## Rebuild contract

```text
.progressstudio v8
    ├── mapping + working tree
    ├── embedded Progress source
    └── embedded BOQ source
            ↓
    restore verified local copies
            ↓
    WorkbookExportService / latest generation engine
            ↓
    latest-format rebuilt workbook
```

Projects created before v8 remain readable. They must be opened with their original/relinked source workbooks once and saved again before they become self-contained. Legacy Actual Progress migration from an edited workbook is intentionally outside MS-R1.

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

### Edited-workbook rebuild boundary (MS-R2)
`.progressstudio` remains the structure/mapping source of truth. An edited exported workbook may be supplied as a secondary, read-only source of user-owned `main` inputs (Activity Amount, weekly Plan, weekly Actual). The migration is applied to the freshly rebuilt `main`, after which all derived workbook views are regenerated. No legacy formulas, formatting, WBS structure, or project-session state are imported.
