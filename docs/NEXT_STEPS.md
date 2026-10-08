# Next Steps — Product Owner notes (not approved implementation)

Recorded after Weekly Compact Final Presentation acceptance, baseline
`main @ 17cc3f3`.

## NS-1 — Rebuild weight-basis conversion

**Question:** A workbook created with Equal or Duration weighting later
receives real Activity Amounts. Can a subsequent Rebuild explicitly change
its weighting basis to Amount while retaining workbook history and user edits?

**Status:** Open question / investigate current implementation first.

Check current Rebuild inputs (the accepted standalone contract rebuilds from
the workbook, not automatically from a new XML), Amount source/provenance,
Activity ID matching, missing/zero/invalid Amount handling, plan baseline and
cumulative-progress recalculation, and effects on Mapping/Payment/EV. No
automatic basis switch has been approved.

## NS-2 — Actual progress preservation on Rebuild

**Question:** If incoming data already carries Actual % Complete, can Rebuild
preserve existing workbook Actual progress rather than overwrite it? How should
conflicts between workbook Actual and incoming Actual be resolved?

**Status:** Open question / investigate current implementation first.

Inspect Activity ID identity, date/period semantics, user-edited Weekly values,
new/removed activities, and any snapshot/live rebuild differences. Do not assume
that incoming XML Actual is currently consumed by Rebuild. No conflict-resolution
or automatic overwrite policy has been approved.

## Boundary

These notes are for future product investigation and decision. They are not
implementation authorization, acceptance criteria, or a change to current
Weight Selection / Rebuild contracts.
