# MS11.1 — Unified Working Tree Model

## Goal

Represent workbook WBS nodes, workbook Activities, user-created WBS nodes, and user-created Activities in one working schedule tree without changing the source workbook.

## Contract

- Every tree node has a stable `node_id` independent of WBS code, Activity ID, name, or display order.
- Workbook node IDs are deterministic and rebuild to the same values.
- User-created node IDs are persisted in the project file.
- The existing Mapping Engine continues to use Activity IDs in MS11.1; later MS11 milestones can migrate mappings to stable node IDs safely.
- The source workbook remains read-only until controlled export.
- Existing project/session versions 1–4 migrate to version 5.

## Scope

- Unified domain model: `WorkingScheduleTree` and `WorkingTreeNode`.
- Original and created nodes exposed through one tree projection.
- Parent/child relationships based on stable IDs.
- Tree validation for orphan nodes, cycles, and duplicate Activity IDs.
- Selection stores stable working-tree identity.
- Existing supplemental UI remains compatible while the editor commands are prepared for MS11.2.

## Out of Scope

- Rename, delete, move, or re-parent workbook-origin nodes.
- Mapping reconciliation after structural edits.
- Rebuilding the Progress workbook hierarchy.
