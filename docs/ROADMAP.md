# Progress Studio roadmap and status

## Accepted V1 baseline

main `d0db778`, tag `progress-studio-v1-stable`.
PO accepted regression reference: 797 passed, 6 skipped, 6 subtests passed, 0 failed.
This is a historical accepted result, not the current working-tree test result.

- MS-1 Equal / Duration: completed and accepted.
- MS-2 XML Amount weighting and Monthly % Complete fix: completed and accepted.
- EV Monetary Sources: completed, including Activity Amount / BOQ Mapping,
  milestone BAC / zero Progress Weight / 0–100 earning.
- MS-3 Planning Overview / Mini Gantt: **CANCELLED** by the Product Owner.
- MS-4 Full Regression Closure: completed; Payment label expectations aligned.

## V1.1 — Product experience (accepted)

Accepted at `main @ b16ddb0`, tag `progress-studio-v1.1-stable`.

## Weekly Compact — accepted

- V1: `main @ c9087ea`, tag `weekly-compact-v1-stable`.
- Final Presentation: `main @ 17cc3f3`, tag `weekly-compact-final-presentation-stable`.
- Windows full regression: 839 passed, 6 skipped, 6 subtests passed, 0 failed.
- Product Owner Desktop Excel acceptance: PASS, including F9, Actual edit/F9,
  cutoff movement, monthly checkpoint collision, WBS outline, Save/Close/Reopen.
- See [Weekly Compact contract](WEEKLY_COMPACT.md).

## Current accepted baseline — XW-1

`main @ c928d59`, tag `xw1-p6-xml-writeback-stable`.
BU-1 and XW-1 are accepted, integrated and closed. The Product Owner records XW-1
Windows automated verification, real NKC1 proof and native P6 compatibility as
passed. This is accepted evidence, not new native testing in XW-2.

## BOQ Update and XML Amount Writeback

The [initiative plan](BOQ_UPDATE_XML_WRITEBACK_PLAN.md) owns accepted directions
and decisions. [P6 Amount writeback](XML_AMOUNT_WRITEBACK.md) owns engine/Desktop
behavior and remaining native acceptance boundaries.

| Milestone | Status |
| --- | --- |
| BU-1 — Amount-only BOQ Update | Accepted and integrated |
| XW-1 — P6 writeback engine proof | Accepted and closed |
| XW-2 — Reviewed desktop export | Implemented for engineering handoff; native Windows UX / PO acceptance / Git integration pending |
| XW-3 / BR-0 | Unauthorized for implementation |

XW-2 adds a Tools command and one review/export dialog using the accepted engine.
No MSP/batch writeback, XML merging or structural BOQ revision is added.

## Deferred investigation (not authorized implementation)

See [Next Steps](NEXT_STEPS.md) for Amount conversion at Rebuild and Actual
progress preservation when incoming data already contains Actual values.

## V1.1 accepted scope

Modern/simple ttk presentation, A layout / C theme / A S-curve icon;
Home and Help flows; target-first Rebuild; ENG/THA on next launch;
Finance/AI navigation hidden; README sheet protection explanation;
product README and documentation organization.

No new workbook engine, custom password, Finance/EV formula redesign or Mini Gantt.
Version metadata remains 2.3.0 until the owner authorizes release numbering.

## Historical planning

The [pre-production roadmap](history/PRE_PRODUCTION_ROADMAP.md) records the old
P0–P11 and WIN tracks. It is not the current release status or an authorization
to resume historical milestones. New milestones require PO authorization.

## Payment Input PI-1

PI-1 is authorized and implemented for engineering handoff against `8156050`
(XW-2 stable baseline). Native Excel and PO acceptance remain pending.
See [Payment Input](PAYMENT_INPUT.md). BP-1/BP-2 BOQ projection and separate
values-only Project snapshot are planning-approved only, not implemented.

PI-1 cumulative WBS correction: fixed full-Amount display with derived sparse states,
zero-only WBS presentation and no COV rows; engineering verification passed (1,159 tests; 6 skipped); Coordinator review pending.
Acceptance remains HOLD; BP-1/BP-2 remain unauthorized.
