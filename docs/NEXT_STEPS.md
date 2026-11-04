# Next Steps — Product Owner notes (not approved implementation)

NS-1 and NS-2 were recorded after Weekly Compact Final Presentation acceptance
at `main @ 17cc3f3`; their investigation status is retained. The new initiative
reference below was reviewed at `main @ 78675f5` on 2026-11-01.

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

## BOQ Update and XML Amount Writeback

Product directions are accepted. BU-1 alone was authorized against `main @ 273636b`
and is implemented for engineering handoff. B1/B2 are accepted for BOQ only.
See the [initiative plan](BOQ_UPDATE_XML_WRITEBACK_PLAN.md) for decisions and
[Roadmap](ROADMAP.md) for status. Next: review the BU-1 engineering package,
complete native Desktop/Excel gates, then obtain PO acceptance and separately
authorized Coordinator integration. XW-1, XW-2 and BR-0 remain unauthorized.
XML monetary/source/milestone/formula/target decisions remain open.

## Boundary

NS-1/NS-2 remain future investigation questions. The initiative above has accepted
product direction with implementation authorization limited to BU-1. These notes are not
implementation authorization, acceptance criteria, or a change to current
Weight Selection / Rebuild contracts.
