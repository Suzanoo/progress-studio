# Create Progress weight selection (MS-1)

## User behavior

- Equal is the default. Each ordinary activity receives 1 dummy unit.
- Duration assigns working hours from P6 PlannedDuration or MSP Duration.
  MSP ISO time hours/minutes/seconds normalize to hours (PT8H30M = 8.5).
  No Start–Finish elapsed time, day-length assumption or Equal fallback is used.
- Explicit source milestones are retained with 0 in both modes. P6 Start/Finish
  Milestone and MSP Milestone are authoritative; same-day dates alone are not.
- WBS/summary nodes receive existing rollup formulas, not additional weights.
- Amount is displayed disabled. API/CLI requests cannot silently use Equal.
- Plan distribution (auto/flat/front/back/bell) remains an independent selection.
- No post-creation weight switch or real XML Amount import is implemented.

## Validation

Duration creation stops before writing the import workbook if an ordinary
activity has missing, malformed, non-finite, zero or negative source duration.
The message lists the Activity IDs/names. An unsupported duration notation is
unusable; a non-time MSP duration such as P1D is not guessed from calendar days.
Milestones always receive zero regardless of source duration. Both modes require
at least one positive ordinary weight and a finite positive total. No existing
output is replaced by these validation failures.

Zero-duration ordinary activities are rejected as unusable, not reclassified as
milestones. If zero-duration ordinary work needs a different product policy,
that needs a separate Product Decision; no alternative is implemented here.

## Workbook lifecycle

- Creation Weight Basis is stored in existing Info key/value metadata.
- Existing internal XML Amount storage carries the computed dummy weight to the
  Amount Mapping engine. This is not an imported cost/custom Amount field.
- Workbook README identifies dummy units; Amount headers retain engine names
  and explanatory comments. Live Dashboard says Total weight, not Project Value.
- Metadata survives both Progress rebuild modes and BOQ Mapping regeneration.
- BOQ mapping can replace current amounts under the existing Mapping contract.
  README then states Current amounts: BOQ Mapping while retaining creation basis.
- Legacy workbooks without this metadata retain their original labeling.
- Number formats, rollups, Plan distribution, calculation/F9 policy, cutoff,
  chart ownership and workbook protection stay with existing helpers/renderers.

## Monetary consumers

Dummy weights are not Contract Value, actual cost or cash flow. Payment progress
continues to use amount-weighted percentages; its copied Amount headers carry
context. Payment-Breakdown Amount headers receive the same explanation through
finalization. Earned Value still requires complete embedded BOQ Mapping; a dummy
Create workbook alone cannot become monetary BAC. Finance continues to require
its own cash/receivable/payable inputs and does not infer cash from dummy weights.
Real source Amount import and field selection remain MS-2.

## Compatibility and validation

Desktop and CLI Create always select a weight basis. `--weight-basis equal` is
the CLI default; `--weight-basis duration` selects source duration. The old
`--amount` / DesktopRunOptions.amount_per_activity parameter is retained only
for call compatibility and does not determine new Create weights. Reader-only
normalization and low-level ImportService calls without weight_basis retain
their legacy schedule-only behavior; production Create supplies weight_basis.

Focused tests: tests/unit/workbook/test_weight_selection.py and
 tests/integration/create_progress/test_weight_basis_lifecycle.py.
Automated package checks do not replace Desktop Excel first-open, F9,
Save/Close/Reopen, Mapping/Rebuild and Product Owner acceptance.
