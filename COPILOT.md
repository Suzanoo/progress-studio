# Copilot Instructions — Progress Studio

Read `README_ROADMAP.md`, the current file under `docs/milestones/`, and the relevant tests before editing code.

## Project purpose

Progress Studio maps BOQ cost items to Primavera Activities and exports a Progress workbook. Correct allocation, recoverability, workbook compatibility, and desktop performance are more important than adding convenience features.

## Non-negotiable rules

1. Do not make Tkinter widgets the source of truth.
2. Keep mapping data in `MappingStore` and domain records.
3. Do not read Treeview rows to calculate, save, or export data.
4. Do not keep Excel workbooks open after loading.
5. Use `read_only=True` and `iter_rows(values_only=True)` for input readers when practical.
6. Do not render every Activity or BOQ item at once; preserve pagination.
7. Do not create thousands of Checkbutton widgets; use the text checkbox columns.
8. Update only affected visible rows after mapping commands.
9. Do not change required workbook column names without an explicit migration plan and tests.
10. BOQ identity must use its stable key/source metadata, never Description alone.
11. Activity identity is `Activity ID`.
12. A BOQ item may be split among Activities, but its total allocation must not exceed 100%.
13. `Amount` and `Allocated` are money; BOQ `Remaining` in the GUI is a percentage.
14. Session files must be written atomically.
15. Validate source workbooks before restoring saved allocations.
16. Relink only when SHA-256 matches; never auto-merge a changed workbook.
17. Any session schema change must increment `SESSION_VERSION`, add a migration, and add migration tests.
18. GUI code must not contain allocation business rules that belong in services/domain code.
19. Every behavior change requires tests.
20. Preserve existing CLI and XML-to-workbook workflows unless the milestone explicitly replaces them.
21. Do not add automatic BOQ detection; the user selects the worksheet.
22. Do not add S-Curve generation or embedded charts to the current V3 roadmap.
23. A Progress workbook is valid only when worksheet `main` exists with the required generated headers. Do not guess renamed worksheet names.
24. Export Activity Amount only into Plan Activity rows in `main`; downstream worksheets must remain formula-driven.
25. Reconcile the total Activity Amount written to `main` with the allocated BOQ amount before replacing the destination workbook.
26. Keep table scrollbars visible and do not replace them with heavy custom widgets or progress bars.
23. Export must read from `MappingStore`, not Treeview rows.
24. Never overwrite the loaded Progress workbook.
25. Preserve the first eleven `BOQ Activity Mapping` columns unless an explicit migration is approved.
26. Export through a temporary workbook and replace the target only after a successful save.
27. Partial mapping export requires explicit user confirmation and a reconciliation summary.
28. Keep the session BOQ key contract stable; use `BOQ ID` as an additional export identity.
29. Never set a worksheet AutoFilter on the same worksheet range as an Excel Table.
30. Validate the saved OOXML package before replacing the export destination.
31. Desktop Excel compatibility takes precedence over openpyxl-only compatibility.

## Git and milestone discipline

- One milestone uses one dedicated branch.
- Start from the previous completed milestone branch.
- Use focused commits for domain/infrastructure, UI, tests, and documentation.
- Keep `.git` in delivered project archives.
- Run the full test suite before delivery.
- `git status` must be clean.
- Update `README_ROADMAP.md`, the milestone document under `docs/milestones/`, and `CHANGELOG.md`.

## MS-5 session contract

A mapping session stores:

- format and schema version;
- saved timestamp;
- Progress workbook path and fingerprint;
- BOQ workbook path and fingerprint;
- selected BOQ worksheet;
- saved path, filename, file metadata, and SHA-256 for each workbook;
- BOQ key, Activity ID, and Share percentage for each allocation.

A mapping session does not store copied workbook rows. On load, the application validates and reads the original workbooks, then restores allocations into `MappingStore`.

## Before completing a change

Run:

```powershell
pytest -q
git status
git log --oneline --decorate -10
```

Confirm the full suite passes and the working tree is clean.


## Session migration and relink rules

- Current session schema is version 2.
- Older supported payloads must pass through ordered migration functions in the repository.
- A future-version payload must fail with a clear message.
- A missing workbook may be relinked through the GUI.
- A relink candidate must match the saved SHA-256 exactly.
- Do not implement automatic reconciliation against edited workbooks without a separately approved milestone and explicit domain rules.
- Loading a session clears Undo history by design because the restored allocations are the new baseline.

## Excel recalculation contract

- `main` is the source of truth for Activity Amount.
- Activity weekly percentage cells are user-entered distribution values; do not rewrite them during mapping export.
- WBS and Project Summary percentages are Excel formulas and must be recalculated after Amount changes.
- Every mapped export must use Automatic calculation, Full Calculation on Load, Force Full Calculation, and `calcId = 0`.
- Do not add hover tooltips, mouse-motion rendering, animation, or continuous Treeview repainting.
- The embedded S-Curve preview is outside the current V3 scope.

## Mapping workspace rules

- Layout preferences may contain only presentation state, never Activity, BOQ, allocation, or session data.
- Full BOQ rows must remain accessible after mapping; do not hide them automatically.
- Do not add row tooltips, hover tracking, animation, flashing, or continuous repaint loops.
- Prefer native PanedWindow, Treeview scrolling, and event-driven updates.
- The generator panel and Workbook Inputs may be collapsed to prioritize mapping space.
