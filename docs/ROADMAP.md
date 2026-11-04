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

## Current accepted baseline — WU-2

BU-1 implementation baseline: `main @ 273636b`; WU-2 stable tag: `wu2-duration-columns-stable`.
WU-2 Plan Duration and Actual Duration is accepted and integrated. See
[User workflow](USER_WORKFLOW.md) for its existing workbook behavior.
This records the accepted baseline; no new regression result is claimed here.

## BOQ Update and XML Amount Writeback — BU-1 engineering handoff

Product directions are accepted: staged BOQ Amount-only updates followed by
structural revisions, and reviewed P6 XML Amount export from one workbook to one
target while preserving the original XML. **Only BU-1 implementation is authorized.**
Native Desktop/Excel and Product Owner acceptance remain pending.

The [initiative plan](BOQ_UPDATE_XML_WRITEBACK_PLAN.md) is the single detailed
source for accepted direction, evidence, proposed gates and unresolved decisions.

| Proposed order | Milestone | Status |
| --- | --- | --- |
| 1 | BU-1 — Amount-only BOQ Update | Implemented for engineering handoff; B1/B2 accepted; native/PO acceptance pending |
| 2 | XW-1 — P6 writeback engine proof | Proposed implementation; XML policy decisions and separate authorization needed |
| 3 | XW-2 — Reviewed desktop export | Proposed; depends on accepted XW-1 |
| 4 | BR-0 — Structural BOQ identity/reconciliation design | Proposed later investigation; structural implementation remains deferred |

BU-1 and XW-1 are technically independent; this order is an engineering
recommendation, not an approved implementation schedule. Completed external
investigations and the one-off recovery/export do not mean these features exist.
BU-1 is not marked accepted or release-ready. MSP/batch writeback and
structural BOQ implementation remain outside the initial proposed scope.

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
