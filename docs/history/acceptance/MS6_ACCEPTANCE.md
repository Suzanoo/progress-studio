# MS-6 Acceptance — Final Workbook Export

MS-6 passes only when all conditions below are true:

1. Export reads Activities, BOQ items, and allocations from `MappingStore`, never Treeview rows.
2. Export validation reports Activity, BOQ, allocation, status, and monetary reconciliation totals.
3. Partial mapping requires explicit user confirmation.
4. The loaded Progress workbook cannot be overwritten.
5. Existing output replacement is explicit.
6. Export uses a temporary file in the destination folder and replaces the final file only after successful save.
7. Existing workbook sheets and formulas remain present.
8. `Amount Mapping` Activity amounts are updated from allocated money.
9. `BOQ Activity Mapping` contains Share %, Allocated Amount, Mapping ID, and BOQ ID.
10. The first eleven mapping columns remain compatible with MS-4.
11. `Mapping Summary` records Complete or Partial status and reconciliation totals.
12. Workbook calculation mode requests full recalculation on open.
13. Stable BOQ export IDs are deterministic for identical source metadata.
14. Automated MS-6 tests and the full test suite pass.
15. `docs/milestones/MS6.md`, `README_ROADMAP.md`, `COPILOT.md`, and `CHANGELOG.md` are updated.
16. Git working tree is clean.


## MS-6.1 hotfix acceptance

- Mapping sheet with allocation rows has one Table-owned AutoFilter and no worksheet AutoFilter.
- Mapping sheet with zero allocations has headers only, with no Table and no AutoFilter.
- Temporary XLSX package validation completes before destination replacement.
- Invalid table relationships, duplicate filter ownership, or malformed table XML fail export safely.
