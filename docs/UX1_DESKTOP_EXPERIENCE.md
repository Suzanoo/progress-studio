# UX-1 desktop experience

Implementation handoff; Product Owner acceptance and Git integration are separate.
This guide supplements existing feature contracts and does not change calculations.

## Navigation and ownership

Home has six entry points: Create, Rebuild, BOQ Mapping, Payment, Earned Value,
and P6 XML Amount export. EV selects the existing EV target in Rebuild; XML opens
its existing reviewed export dialog. A running Rebuild retains its current target.
Finance and AI remain hidden. Settings and the header expose the same ENG/THA
preference, applied on next launch. Help includes keyboard and source guidance.

Each workspace retains its own sources. Creating a workbook displays its path
and the next Excel step; it does not replace the current Mapping workspace.
Choose the exported workbook explicitly when starting another task.

## Mapping

Progress WBS/Activities stay on the left, BOQ WBS/Items on the right. BOQ group
rows represent WBS paths; only existing BOQ leaf identities enter allocations.
Filtering and pagination remain store-owned; a WBS may therefore appear on more
than one page. Group expansion is a view operation, not an allocation operation.

Both Trees have independent horizontal and vertical scrolling. Column contents
retain their widths instead of becoming unreadable when the window narrows.
Drag the divider, or Tab to it and use Left/Right. The application never disables
Map/Unmap based on width. F11 provides Focus Mapping; Escape exits when focus is
outside an editor. Workbook Inputs can be folded. When a resized, loaded workspace
has less than four Tree rows of height, Inputs fold without changing the saved
preference. Users can reopen them. At very small heights, enlarge the window or
fold Inputs to make more rows visible. 800x640 Focus Mapping was exercised on
Linux Tk; Windows scaling and font metrics remain acceptance gates.

Structure contains Add/Edit/Delete, Arrange, and Tree View menus. Mapping Undo
and structure Undo/Redo remain separate histories. The allocation strip shows
the target Activity, checked BOQ count, count outside the current page and Share.
Off-page means pagination/filter exclusion; collapsed descendants remain on-page.
Successful Map/Unmap retains the Activity and clears BOQ checks as before.

| Focus / action | Behavior |
| --- | --- |
| Tree arrows / Home / End / Page Up / Page Down | Native focus navigation |
| Space on Activity | Select the one mapping target |
| Space on Progress WBS | Select the structural WBS, without creating an allocation |
| Space or Ctrl+Space on BOQ leaf | Toggle that leaf additively |
| Shift+Up/Down on BOQ leaf | Extend checked range using the existing filtered sequence |
| Ctrl+A on BOQ Tree | Add the current page; selection menu also offers all filtered |
| Ctrl+C on Tree | Copy selected stable IDs, names and BOQ amounts as tab-separated text |
| Ctrl+F on Tree | Focus that pane's search |
| Right-click / Shift+F10 | Context commands; do not replace checked business selection |
| Delete with Mapping Tree focus | Existing Unmap operation; never structural deletion |
| Alt+arrows on Progress Tree | Existing structure movement commands |
| Entry/Text/Spinbox/Combobox | Toolkit editing, clipboard and selection retain ownership |

Project commands are owned by the shell and active Mapping workspace. A menu
command invoked elsewhere first navigates to Mapping; repeat there to execute.
Text editing and modal dialogs do not trigger Mapping Undo/Delete. Tree paste
is not a new business import operation. Wheel handlers are local to each page
or Tree and do not remove another workspace's bindings.

## Persistence and operation feedback

Mapping save failure leaves in-memory changes marked unsaved, shows the error,
and suppresses the saved-success notification. Use Mapping > Save to retry, or File > Save As (Ctrl+Shift+S).
Closing or replacing a dirty Mapping workspace requires an explicit discard
choice. Project save remains the existing atomic repository operation. Failure
to update recent-project history does not falsely invalidate a completed save.

Payment typing or Paste invalidates the prior validated source immediately.
Press Enter or Check source; Browse validates the selected file. Prepare and
Breakdown require the displayed and validated paths to agree. Source and period
controls remain disabled during work. Worker outcomes are consumed on the Tk
thread, including failures, so input and retry controls can recover reliably.

Application close inspects all live workspaces and child dialogs for active
operations, including BOQ preparation/export and Word/payment reports. It keeps
Tk polling alive until completion. XML retains its existing cancellation request;
other cancellation belongs to the owning dialog. It does not kill worker threads.

## Presentation

The Fieldwork palette uses ivory surfaces, graphite navigation, orange primary
actions and blue selection. Existing S-curve branding is retained. Home, Create,
Payment, Rebuild, Settings and Help use scoped scrolling; focus reveals controls
when reached with Tab. Create's Activity Log is expandable; its result includes
an explicit next step. Payment's independent report actions remain independent
and use a horizontal workflow row, with scrolling and focus reveal in narrow windows.

Existing XML, BOQ and report dialogs retain their review, diagnostics, cancellation
and output-publication services and inherit shared styling/native text controls.
Workbook formulas, templates, source-reconciliation rules and monetary sources
are not changed by this desktop work.

## UX-2 refinement

The persistent top Mapping toolbar is removed. The native Mapping menu contains
Open, Save and Export Mapped Workbook; existing File/keyboard commands remain.
The single action bar below the Trees is Share %, Map, Unmap, Undo, Clear all.
By explicit PO decision for UX-2, **Clear all clears checked Activity/BOQ rows
only**, including off-page checks; it does not clear allocations or Undo history.
The previous destructive operation remains in the Tree context menu as
**Clear all mappings…**, with the same confirmation, persistence and Undo.

Workbook Inputs occupies one compact collapsible header, including Update BOQ
Amounts. BOQ WBS-2/3/4 filters use actual stored hierarchy fields. Changing a
parent resets descendant filters, preserving checked leaf identities. WBS-4 is
unavailable when no matching child level exists. Long values are stored intact.
Filters use wider dropdowns in one row. If the pane is narrow, a horizontal
scrollbar appears; Tab reveals each focused control. Trees remain independently
scrollable. No width-dependent Mapping action restriction is introduced.

The Payment BOQ workflow is horizontal. Save in Excel remains informational
guidance, not an executable application command. Prepare, Excel export and Word
reports retain independent handlers and eligibility. Scroll/Tab reveal provides
access on narrower windows; the row does not impose a new workflow dependency.
Small monochrome Pillow/Tk command images have strong owner references and
text-only fallback. Native menus retain text for Windows reliability.
