# Final Workbook Policy Hotfix

This hotfix closes the workbook finalization contract after FINISH-1.

## Protection

- Every final worksheet is protected.
- Intended user input/control cells remain unlocked.
- Workbook structure is protected to prevent accidental sheet rename/delete/move/insert in Excel.
- Progress Studio Rebuild remains able to replace sheets because the protection is an Excel UI guard, not file encryption.
- Internal protection password is centralized in `progress_studio/config/workbook_protection.py`.

## F9 / Save

All final portable workbooks use Manual calculation with `calcOnSave=true`.

- **F9 / Save:** recalculates Excel formulas.
- **Rebuild:** regenerates Python-owned snapshots/caches and generated views.

F9 / Save is intentionally not described as a Python rebuild mechanism.

## Rebuild ownership

Progress and Payment rebuild paths apply the final workbook policy once at the output boundary. Payment-only rebuild no longer reapplies protection/visibility after the renderer has already finalized the output. Live Payment renders the Payment sheet first and finalizes afterward, so the newly-created sheet is protected too.
