# MS-7 — Mapping Workspace UX

## Goal

Give BOQ mapping the largest practical screen area while keeping the Tkinter interface lightweight and predictable.

## Delivered

- Start the desktop window maximized while preserving normal Restore behavior.
- Add `Focus Mapping` to collapse the Primavera generator panel and let the mapping workspace use the full application width.
- Add `Show Generator` to restore the generator panel without reopening the application.
- Make Workbook Inputs collapsible and automatically collapse them after both workbooks and the BOQ worksheet are ready.
- Keep Activity and BOQ tables inside a user-adjustable horizontal PanedWindow.
- Persist the generator collapsed state, Workbook Inputs collapsed state, and mapping divider position in a small user-level JSON file.
- Compact session, export, mapping, and BOQ filter controls.
- Increase Description column space while reducing low-value WBS column width.
- Preserve Full BOQ rows; no automatic hiding after mapping.
- Add no tooltips, hover effects, animation, progress bars, embedded charts, or continuous repaint behavior.

## Preference file

Presentation preferences are stored separately from project data:

```text
~/.progress_studio/layout.json
```

Failure to read or write this optional file never blocks mapping.

## Branch

`feat/ms7-mapping-workspace-ux`
