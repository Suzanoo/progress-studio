# Progress Studio V3 Roadmap

Progress Studio is a desktop workflow that connects Primavera schedule data with BOQ cost data, allocates BOQ amounts to Activities, exports a mapped Progress workbook, and generates reporting outputs.

## End-to-end workflow

```text
Primavera XML
    ↓
Progress workbook
    ↓
Load Progress workbook + BOQ workbook
    ↓
Map BOQ items to Activities
    ↓
Allocate each BOQ item by percentage
    ↓
Save and resume the mapping session
    ↓
Export mapped Progress workbook
    ↓
Generate S-Curve on demand
```

## Milestones

### MS-1 — Independent Workbook Loader

**Goal:** Mapping must not require regenerating Primavera XML every time.

Completed:

- Load an existing Progress workbook.
- Load an existing BOQ workbook.
- Let the user select the BOQ worksheet manually.
- Validate required worksheet columns.
- Read Excel with streaming readers.
- Close workbooks after reading.
- Store loaded Activities and BOQ items in memory.

Branch: `feat/ms1-independent-loader`

### MS-2 — High-performance Mapping UI

**Goal:** Keep the desktop UI responsive with large BOQ worksheets.

Completed:

- Paginated Activity and BOQ tables.
- Lightweight text checkboxes (`☐` / `☑`).
- Activity WBS hierarchy headers.
- Activity and BOQ search.
- Indexed WBS-2 and WBS-3 filters.
- Summary of loaded and mapped data.
- Treeview renders only visible pages, not the complete dataset.

Branch: `feat/ms2-mapping-ui`

### MS-3 — Mapping Engine

**Goal:** Implement reliable BOQ-to-Activity mapping.

Completed:

- Select exactly one Activity.
- Select one or more BOQ items.
- Map, unmap, and undo.
- Update Activity Amount immediately.
- Calculate BOQ Allocated, Remaining, and Status.
- Update only affected visible rows.
- Keep the Mapping Store as the source of truth.

Branch: `feat/ms3-mapping-engine`

### MS-4 — Share Allocation

**Goal:** Allow one BOQ item to be split across multiple Activities.

Completed:

- User-entered Share percentage.
- Multiple Activity allocations per BOQ item.
- Combined allocation cannot exceed 100%.
- Re-mapping the same BOQ/Activity pair replaces its previous share.
- Activity Amount is calculated from allocated money.
- BOQ table shows Amount and Allocated as money.
- BOQ Remaining is shown as a percentage.
- Status values: Unmapped, Partial, Full.
- Export structure includes Share and Allocated Amount.

Branch: `feat/ms4-share-allocation`

### MS-5 — Persistent Mapping Session

**Goal:** Let the user stop mapping and continue later without losing work.

Completed:

- Save mapping session as `*.mapping.json` or another selected JSON filename.
- Load a session and automatically restore both workbooks, BOQ worksheet, and allocations.
- Auto-save after Map, Unmap, Undo, and Clear All once a session file exists.
- Atomic session writes through temporary-file replacement.
- SHA-256 workbook validation before restoring a session.
- Reject changed, missing, malformed, or incompatible session inputs.
- Recent Sessions list.
- Clear All with confirmation and immediate Undo recovery.
- Session status shown in the mapping toolbar.

Branch: `feat/ms5-persistent-session`

### MS-6 — Final Workbook Export

**Goal:** Produce the final mapped workbook safely from the in-memory mapping state.

Completed:

- Added export validation and a reconciliation summary.
- Added explicit confirmation for partial mapping exports.
- Updated `Amount Mapping` Activity amounts and statuses.
- Rebuilt `BOQ Activity Mapping` as a formatted Excel table.
- Added `Mapping Summary` as the first worksheet.
- Preserved existing worksheets and formulas.
- Added stable BOQ export IDs while preserving the existing session key contract.
- Prevented overwriting the loaded source workbook.
- Added atomic temporary-file export and explicit overwrite policy.
- Requested full Excel recalculation on open.

Branch: `feat/ms6-workbook-export`

### MS-7 — On-demand S-Curve Generator

**Goal:** Generate the S-Curve only after real mapped amounts are available.

Planned:

- Remove the permanently embedded S-Curve from the mapping workspace.
- Add a `Generate S-Curve` action.
- Generate weekly and monthly views from the current mapped workbook.
- Allow generation from partial mapping only after explicit confirmation.
- Keep chart generation separate from the mapping UI.

Planned branch: `feat/ms7-scurve-generator`

### MS-8 — Production Polish

**Goal:** Prepare Progress Studio for routine project use.

Planned:

- Final error handling and user messages.
- Keyboard shortcuts and practical context actions.
- Loading indicators for slow file operations.
- Performance profiling with large workbooks.
- Logging and diagnostic export.
- Final release notes, packaging, and acceptance tests.

Planned branch: `feat/ms8-production-polish`

## Current status

| Milestone | Status |
|---|---|
| MS-1 Independent Workbook Loader | Completed |
| MS-2 High-performance Mapping UI | Completed |
| MS-3 Mapping Engine | Completed |
| MS-4 Share Allocation | Completed |
| MS-5 Persistent Mapping Session | Completed and hardened |
| MS-6 Final Workbook Export | Completed |
| MS-7 On-demand S-Curve Generator | Planned |
| MS-8 Production Polish | Planned |

## Architecture rules

```text
Excel readers
    ↓
Domain records
    ↓
MappingStore / session state
    ↓
Application services
    ↓
Tkinter presentation
```

- GUI widgets never own business data.
- Excel files are not kept open after loading.
- Mapping operations do not read Excel again.
- The Mapping Store is the in-memory source of truth.
- Session JSON is persistence, not the runtime data model.
- Session files support explicit schema migrations.
- Relink accepts only moved or renamed workbooks with identical SHA-256 content.
- Changed workbooks are never reconciled or merged automatically.
- Export reads from domain records and allocation records, not Treeview rows.
- Large tables must use pagination and update only affected rows.
