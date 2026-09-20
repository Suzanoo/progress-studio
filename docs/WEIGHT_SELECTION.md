# Create Progress weight selection (MS-1 / MS-2)

## User behavior

- Equal is the default. Each ordinary activity receives 1 dummy unit.
- Duration assigns working hours from P6 PlannedDuration or MSP Duration.
  MSP ISO time hours/minutes/seconds normalize to hours (PT8H30M = 8.5).
  No Start–Finish elapsed time, day-length assumption or Equal fallback is used.
- Explicit source milestones are retained with 0 in both modes. P6 Start/Finish
  Milestone and MSP Milestone are authoritative; same-day dates alone are not.
- WBS/summary nodes receive existing rollup formulas, not additional weights.
- Amount requires explicit selection of a declared numeric custom XML field.
  The picker displays name/alias, source-qualified identity/type, raw values,
  ordinary positive/zero counts, milestone count, total and validation issues.
  No field is selected automatically, even if only one exists or its name says Amount.
  Changing the XML path clears selection; editing its contents requires preview again.
- Plan distribution (auto/flat/front/back/bell) remains an independent selection.
- No post-creation weight switch is implemented. Native schedule costs are untouched.

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

### Amount validation and supported source representation

P6: declared Activity UDF Double, Integer or Cost, keyed by TypeObjectId and
read from the corresponding DoubleValue, IntegerValue or CostValue.
MSP: declared Number/Cost custom fields (numeric CFType / native NumberN or
CostN field names), keyed by FieldID, with explicit task ExtendedAttribute Value.
Display aliases and names do not determine identity. Numeric-looking Text/date/
duration/flag fields are not treated as Amount. Missing declarations are not
guessed from IDs; lookup-only values without explicit numeric Value fail validation.

Ordinary activities require a selected field value that is numeric, finite,
non-negative and representable by the existing workbook float model. Missing,
blank, nonnumeric, negative, duplicate/ambiguous field values, unsupported numeric
range and a nonpositive/unusable ordinary total stop creation before writing.
Errors identify Activity ID/name and field. Individual zero Amount is valid and
shown in preview. Explicit milestones remain zero even if their field is missing
or contains a nonzero value; preview states that the source value is ignored.

Raw decimal strings are retained until validation. Preview totals use Decimal;
conversion to the existing float/Excel numeric model occurs once for weights,
without rounding amounts to cents. Display uses two decimals. Excel's numeric
precision remains a limitation; raw values are shown for review before Create.

## Workbook lifecycle

- Creation Weight Basis is stored in existing Info key/value metadata.
- Existing internal XML Amount storage carries the computed dummy weight to the
  Amount Mapping engine. In Amount mode it carries the explicitly selected
  monetary XML values; in Equal/Duration it carries only computed dummy weights.
- Workbook README identifies dummy units; Amount headers retain engine names
  and explanatory comments. Live Dashboard says Total weight for dummy mode and
  Project Value for monetary mode. Amount stores field identity/name/source/type/
  native name in Info and identifies the selected source on README.
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
Amount Create does not bypass EV's BOQ requirement: complete BOQ Mapping is
still needed before EV generation. No synthetic BOQ/allocation is created from
XML values. After Mapping, BOQ totals become current monetary authority; XML
field metadata remains creation provenance. Finance still requires independent
cash/receivable/payable inputs even when XML Amount is monetary.

## Compatibility and validation

Desktop and CLI Create always select a weight basis. `--weight-basis equal` is
the CLI default; `--weight-basis duration` selects source duration.
`--weight-basis amount --amount-field p6:udf:<ObjectId>` or
`--weight-basis amount --amount-field msp:field:<FieldID>` selects the exact field
shown in the GUI picker. No identity is hardcoded for a customer/sample. The old
`--amount` / DesktopRunOptions.amount_per_activity parameter is retained only
for call compatibility and does not determine new Create weights. Reader-only
normalization and low-level ImportService calls without weight_basis retain
their legacy schedule-only behavior; production Create supplies weight_basis.

Focused tests: tests/unit/workbook/test_weight_selection.py and
 tests/integration/create_progress/test_weight_basis_lifecycle.py.
Automated package checks do not replace Desktop Excel first-open, F9,
Save/Close/Reopen, Mapping/Rebuild and Product Owner acceptance.

MS-2 tests: tests/unit/xml/test_amount_field_selection.py,
tests/unit/workbook/test_amount_metadata.py,
tests/integration/create_progress/test_xml_amount_workflow.py,
tests/integration/desktop/test_amount_selection_gate.py, and
tests/regression/weights/test_amount_monetary_lifecycle.py.

MS-1 Desktop Excel acceptance remains accepted. MS-2 requires its own Desktop
Excel/application acceptance; automated checks do not claim that gate.
Payment Label width 150 px is accepted by PO; its two old test expectations
remain unchanged and maintenance is deferred until after MS-4.
