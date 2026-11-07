# Payment Input: Amount and cumulative WBS requirements

PI-1 is implemented for engineering review against `main @ 8156050`.
Native Windows Excel acceptance and Product Owner acceptance remain pending.
BP-1 is accepted. BP-2 export is a separate derived workbook; see [BOQ Payment Export](BOQ_PAYMENT_EXPORT.md).

## Prepare and edit

Payment → Prepare Payment Input creates or reconciles the embedded sheet:

`Type | WBS | Activity ID | Activity Name | Amount | P01 | P02 | ...`

Amount is a locked snapshot of current `main` Activity/P Amount, matched by exact
Activity ID. Prepare refreshes it. Editing main and pressing F9 alone does not
refresh the snapshot. Payment Rebuild continues to preserve Payment Input.
Equal, Duration and Amount values are weights; they are not necessarily BOQ money.
Zero remains numeric zero. Actual rows are excluded, including duplicate Actual IDs.

Numeric saved results of formula Amounts are supported by file-based Prepare and
standalone creation. Recalculate and save the source workbook first. A cached
result's freshness cannot be guaranteed. In-memory reconciliation only uses a
cached result when explicitly supplied with matching Activity ID and formula;
otherwise it displays `Formula unavailable`. Missing and invalid values display
`Missing Amount` and `Invalid Amount`. Text numbers, booleans, negative values and
non-finite values are not silently coerced into valid weights.

Activity period cells are editable. Legacy standalone Payment Dates in row 7
also remain editable. Amount and WBS summaries are locked.
Freeze panes retains identity and Amount at the left. Existing files with P01 in
column E remain readable; preparing them upgrades the presentation. Period IDs
and requirements are preserved by Activity ID and period ID, including nonsequential
IDs. Removing a populated period is rejected rather than discarding its inputs.
New Activities retain the existing suggestions; deleted Activities are counted
by existing reconciliation. Malformed existing requirements stop Prepare.

## Cumulative WBS Requirement

PI-1 correction supersedes the original specified-only average and COV rows.
WBS summaries now use a fixed denominator: all descendant Activity Amount.
For each period, multiply every Activity Amount by its derived cumulative
Requirement state and divide by the full WBS Amount. Never sum cumulative
percentages across periods, and never count nested WBS or Actual rows again.
Hierarchy follows outline position, so repeated names remain separate branches.

Derived states are **display only**:

- An entered valid value is used exactly, including explicit zero.
- Before the first entry, use zero if the Activity has at least one valid entry
  somewhere in its period schedule.
- After an entry, a blank uses the latest entered value.
- A positive-weight Activity with no entries anywhere remains unknown. A partly
  specified WBS shows `Incomplete`, not a partial final percentage.
- A wholly blank branch has a blank summary. An empty branch remains blank.
- Missing/invalid weights show `Check Amount`; malformed requirements show
  `Check input`. These checks take priority over blank/no-weight states.
- A zero-total branch with entered requirements shows `No weight`.
- Zero-weight milestones do not affect the weighted result or completeness,
  but retain their original Payment eligibility and input diagnostics.

Valid WBS numerical zero is hidden by a zero-specific number-format section;
it remains numeric zero. Small nonzero values remain displayed even if they
round to 0.0%. Explicit Activity zero remains visible, with its original format.

Actual decreasing inputs remain unchanged and may lower WBS results. Amber
conditional formatting identifies the offending Activity cell and affected WBS
periods. The warning persists in later WBS periods until the contradiction is
corrected. The fixed diagnostic line identifies the first affected Activity ID;
Amount/input/completeness issues take priority there. No new eligibility gate,
clamping, sorting or cumulative maximum is introduced.

The schedule is an editable target schedule, not a history of what was known
at each date. Adding/removing an Activity's only entry can therefore change
previous summary periods between unknown and numeric states. Changing Amount
or hierarchy during Prepare can also change prior summaries.

The frozen header describes the calculation and derived blank rules. COV rows
are removed during Prepare/Reconcile; diagnostics use hidden locked helpers
and one fixed header line. No new analysis sheet is introduced. Old Input is
not silently restructured by opening it or by Payment Rebuild. Run Prepare once
to migrate the old presentation while preserving requirements by ID and period.

F9 / Save recalculates the summary using the existing Manual/calculate-on-save
policy. Python does not evaluate formulas or populate caches. Fresh outputs may
show blank formula results until Excel recalculates. Amount is still a Prepare-
time snapshot. These weights may be Equal, Duration or Monetary; the summary is
not Actual progress, verified BOQ money, approved payment or an incremental bill.

## Ownership and compatibility

The embedded and existing standalone creation paths share summary presentation
and Amount validation. Standalone Payment Date inputs retain their previous role;
embedded row 7 remains reserved. No Payment Position Engine, cumulative Plan,
Mapping share, BU-1, EV, XML or Rebuild calculation rule changes.

Prepare validates before saving, stages output in the destination directory,
reuses opaque package preservation, validates tables and atomically publishes.
Failure before publication leaves the original and any existing output intact.

## BOQ Payment projection boundaries (BP-1/BP-2)

Only explicit Mapping Shares establish eligible BOQ-to-Activity relationships.
Activity Amount alone never implies mapping, including monetary Activities such
as systems work that the user intentionally leaves unmapped.

- Fully mapped BOQ items with complete required inputs may have final percentages.
- Partial items retain known contributions and coverage; no misleading final %.
- Unmapped BOQ items remain in a future full Project snapshot with blank Payment
  percentages and `Unmapped` status.
- Preserve original BOQ files. Export a separate values-only snapshot; editable
  Payment Input remains embedded in Progress.
- Do not re-import derived Payment reports into BU-1 automatically.

The future projection must define and validate complete source identities,
compatible Mapping revisions and sparse input coverage. It is not part of PI-1.

### Consecutive completion display

WBS cells retain their cumulative numeric result in every period. A conditional
number format hides a cell only when both it and the immediately preceding
period are numeric and exactly 1 (100%). The first period is never subject to
this rule. A decrease remains visible; a subsequent return to 100% is shown.
Values merely rounding to 100.0% are not completion. Existing exact-zero
suppression, diagnostic text and amber decrease warnings remain intact.
Activity formatting and sparse inputs are unchanged. F9 recalculates the display.
Existing workbooks receive this presentation during supported Prepare/Reconcile;
opening or Rebuild alone does not silently migrate their layout or formatting.

### Overall Project Summary

Payment Input starts with one locked, level-zero overall summary. It uses the
`main` Plan-side Project Summary description and WBS identifier when present;
its Activity ID is always empty. For older workbooks without project metadata,
an existing WBS enclosing the complete tree is retained as the root. Otherwise
an explicitly labelled `Project Summary` with no invented WBS/Activity ID is
prepended. Level-zero legacy branches shift uniformly to become children;
existing positive hierarchy levels are preserved.

The overall Amount and cumulative Requirement use the same leaf Activity-only
calculation and diagnostics as other WBS summaries. Nested WBS totals and Actual
rows never enter the denominator. Zero suppression, consecutive-100 suppression
and decrease warnings apply unchanged. Prepare/Reconcile installs this layout
and preserves saved sparse entries by Activity ID/period. Rebuild does not
silently migrate existing layouts. Embedded and standalone paths are supported.
