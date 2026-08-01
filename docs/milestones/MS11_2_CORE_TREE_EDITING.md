# MS11.2 — Core Working Tree Editing

Progress Studio now treats the imported schedule as an editable working tree while preserving the source workbook.

## Scope

- Rename original or user-created WBS nodes.
- Rename original or user-created Activities, including Activity ID validation.
- Delete a WBS subtree or Activity from the working tree after mapped BOQ allocations are removed.
- Move nodes up or down within the same parent.
- Re-parent WBS nodes and Activities beneath another WBS.
- Undo and redo structural tree edits.
- Persist the complete working-tree snapshot in project format version 6.
- Migrate project versions 1–5 to version 6.

## Safety contract

- The source Progress workbook is never overwritten by tree editing.
- Activity IDs remain unique.
- WBS parent cycles are rejected.
- Activities cannot be moved to the root.
- Deleting mapped Activities or WBS subtrees is blocked until allocations are removed.
- Renaming an Activity ID keeps its BOQ allocation references aligned.

Workbook reconstruction and export of edited original nodes remain part of a later milestone.
