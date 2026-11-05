# Progress Studio User Workflow

This is the current high-level workflow. Historical v2.3 guides are
archived under `docs/history/user-guides-v2.3/` and should not be
treated as the current contract.

## 1. Create Progress

Input:

-   Microsoft Project XML, or
-   Primavera P6 XML.

Progress Studio normalizes the schedule and creates the initial Excel
workbook.

The initial workbook includes weekly `main`, monthly `main_monthly`,
Dashboard and internal/helper data required by the selected workflow.

Choose **Equal** (default), **Duration**, or **Amount** as the Weight basis.
For Amount, choose **Use XML field** or **Enter in Excel later**.
For XML, choose **Select Amount field / Preview...**, select the intended
numeric custom field, review its raw values/counts/total, and confirm the selection.
No field is chosen automatically. No fallback amount needs to be entered.

**Enter in Excel later** works with or without XML Amount fields. It starts ordinary
Activities at 1 and milestones at 0. These are Progress weights, not EV Contract
Values. Edit `main` Amounts when useful; there is no monetary confirmation step.
After Amount edits, F9 / Save and **Progress Rebuild** refresh generated/copied
Monthly and Dashboard values. F9 alone does not rewrite those copies. Payment-only
or EV-only Rebuild does not replace this Progress refresh.
Create includes blank `EV Monetary Inputs` for this path, but no generated Payment
or Earned Value views. To use Activity-source EV, enter independent Contract Values
there; changing `main` Amounts does not populate them automatically.

Equal gives each ordinary activity one calculation unit. Duration uses the
source working duration normalized to hours (P6 PlannedDuration / MSP Duration),
not elapsed Start–Finish time. Milestones stay in the workbook with zero weight;
summary/WBS rows roll up activity weights. Invalid ordinary duration stops
Duration creation with the affected Activity IDs; it never switches to Equal.

Equal/Duration dummy weights are not Contract Value, actual cost or cash flow.
Amount with **Use XML field** uses monetary values from the selected XML field; this does not
supply actual cost or Finance cash data. Missing, invalid or negative ordinary
Amount and an unusable total block creation; individual zeros are retained. The workbook
README identifies the basis; Info preserves it through Save/Reopen and Rebuild.
BOQ Mapping still owns real amount allocation and updates the current-basis label.
Weight is independent of the existing auto/flat/front/back/bell Plan distribution.
See [Weight selection](WEIGHT_SELECTION.md) for validation and compatibility details.

## 2. Mapping (optional)

Load the Progress workbook and BOQ workbook, then allocate BOQ amounts
to Activities.

Activity identity is Activity ID. Mapping owns BOQ -\> Activity
allocation; it does not own Payment or Progress rebuild behavior.

## 3. Payment (optional)

The Payment Workspace contains two separate workflows.

### 3.1 Standard Payment

`Payment Input` stores persistent payment requirements. Standard Payment
rendering uses the current workbook schedule/progress structure to
position `Payment` output.

Standard Payment identity is Activity ID.

Valid Plan profiles and Payment Input requirements are sufficient: Standard Payment
does not require real Amounts, BOQ Mapping or replacement of initial weights.
Requirements must be finite percentages from 0% to 100%; blank is not explicit 0%.

Payment must not rebuild Progress-owned Dashboard/overlay data.

### 3.2 Payment Breakdown

Payment Breakdown derives consolidated progress directly from current
`main` and creates/replaces `Payment-Breakdown`.

Its grouping identity is the **exact complete Activity Name after
trimming leading/trailing whitespace**.

Payment Breakdown does not automatically use fuzzy matching, contains
matching, keyword regrouping, case normalization, or Activity ID
grouping. If names should match but do not, correct the Activity Names
in `main` and build Payment Breakdown again.

For repeated exact names, each source Activity keeps its own progress
profile and the combined period progress is weighted by Activity Amount.
Combined cumulative progress is then calculated from that weighted
profile.

The output keeps source Activity rows as collapsible Excel detail rows
while Combined Progress and Combined Cumulative remain visible as the
group summary.

Payment Breakdown is **not** part of Payment Rebuild. After changing
relevant Activity names, Amounts, or progress in `main`, run **Build
Payment Breakdown** again from the Payment Workspace.

See [Payment Breakdown](PAYMENT_BREAKDOWN.md) for the detailed feature
contract.

## 4. V1.1 navigation and help

Home shows Create/Rebuild as primary actions and Mapping/Payment as optional.
Home's View guide and the persistent Help entry open the same first-time/cycle guide.
Select ENG or THA at the top right. The saved preference applies on the next launch;
current work is not restarted. Workbook formulas, sheet names and identities stay unchanged.

Finance and AI Helper are hidden from normal navigation. Finance's underlying
implementation is retained; see [retained Finance reference](FF3_DESKTOP.md).

## 5. Work in Excel

After initial generation, `main` is the workbook source of truth.

Typical edits include Actual progress and other approved editable
inputs. Structural edits such as adding/deleting Activities, WBS or
timescale periods require Rebuild to regenerate Python-owned derived
outputs.

### Timescale labels

``` text
X  X | W1 W2 ... Wn | X X
X    | M1 M2 ... Mn | X
```

-   `X` = display margin.
-   `Wn` / `Mn` = reporting-period labels.

## Weekly Compact presentation

Weekly Compact is an additional presentation of the existing Weekly model,
not a new Monthly calculation model. The Gantt body and four-row monthly footer
follow `main_monthly`; the S-Curve keeps Weekly points and shows monthly
checkpoints plus purple current-cutoff Plan/Actual markers. The editable Weekly
cutoff remains on `main`; the Compact cutoff date is a locked linked display.
Monthly table values are not masked to the Weekly cutoff. Press F9 after opening
or editing to update Excel-calculated presentation, then Save. See
[Weekly Compact](WEEKLY_COMPACT.md).

The Weekly and Monthly cutoff dropdowns are in B:C below each generated table
and S-Curve footer. Weekly Compact shows the Weekly cutoff there as a read-only
date. Dashboard and EV reporting dates remain separate. Use Progress Rebuild
to migrate an older workbook's controls; Payment/EV-only rebuild does not move
them. Regeneration retains the existing cutoff default/reset behavior, so check
the selected dates after Rebuild. If placement reports an occupied or merged
cell, preserve the original workbook, move the conflicting content to a free
location in a copy, and retry. Rebuild will not overwrite that content silently.

## 6. Recalculation in Excel

### Calendar duration columns (WU-2)

New Create and fresh Mapping-generated workbooks include **Plan Duration** before
**Plan Start** and **Actual Duration** before **Actual Start**. Values are inclusive
calendar days (`INT(Finish)-INT(Start)+1`), displayed with `d`, not source working
hours and not Duration weighting. Ordinary same-day activities show 1; explicitly
identified milestones show 0 when dates are valid. Amount=0 never identifies a
milestone. Imported milestone identity is retained independently of monetary data.

Plan Duration is on the P row. Actual Duration is on the A row only when both
Actual Start and Actual Finish are present and valid (completion is established
by Actual Finish, not progress percentages or a cutoff). Missing, nonnumeric or
reversed dates produce a blank value; entered invalid/reversed date pairs mark
the duration cell red. Time-of-day is ignored. Plan summaries use the span of
their existing rolled-up dates, not a sum of durations; Actual summaries stay blank.
This does not introduce a new automatic rollup of edited child dates.

Edit dates in `main`, then F9/Save; duration cells are protected derived values.
Monthly metadata dates remain copies under the existing contract: Progress Rebuild
refreshes them and their durations. Live Compact's metadata follows `main`;
Snapshot Monthly/Compact contain durations computed at Rebuild. Formula-based
date inputs need valid saved Excel caches for Snapshot calculation.

Ordinary Rebuild does not insert these columns into legacy `main`. Fresh Mapping
generation can use the new schema; unmarked legacy activities are not guessed to
be milestones. Mapping reads paired Actual dates from its loaded source workbook.
The optional edited-workbook migration retains its existing selective progress/
Amount contract; it does not newly migrate schedule dates. Load the intended
schedule-date source into Mapping. No automatic legacy column migration is included.

Progress Studio workbooks use a user-driven Excel formula calculation
policy.

-   **F9** recalculates Excel formulas.
-   **Save** asks Excel to calculate formulas on save.
-   Neither action runs Python code.

If a generated snapshot/cache/view must be regenerated after structural
changes, use Progress Studio Rebuild.

## 7. Rebuild

Rebuild uses the selected workbook itself as the source.

The workspace exposes Snapshot/Live workbook modes and Progress/Payment
scope.

``` text
                    Progress    Payment
Snapshot               ✓           ✓
Live                   ✓           ✓
```

### Progress Rebuild

-   preserves `main` as the edited source;
-   rebuilds Progress-owned derived outputs for the selected mode;
-   preserves Payment/user-owned data according to the contract;
-   does not renumber the weekly W/X labels in `main`.

### Payment Rebuild

-   rebuilds `Payment` only from current `main + Payment Input`;
-   preserves Progress-generated views;
-   does not own `Payment-Breakdown`.

## When to Rebuild

Use Rebuild when the workbook structure or Python-owned generated data
needs regeneration, for example after:

-   adding/deleting Activities;
-   changing WBS structure;
-   extending/shrinking the project timescale;
-   major Plan/rebaseline changes;
-   changing Payment Input that requires standard Payment to be
    regenerated.

Use **Build Payment Breakdown** instead of Rebuild when only the
`Payment-Breakdown` output needs to be refreshed.

## V1.1 Rebuild presentation

Select the saved workbook, then Target: Progress, Payment or Earned Value.
Progress/Payment show their existing Snapshot/Live options. EV shows Live monetary-source
controls and preparation when applicable. Ordinary EV progress edits need F9/Save;
structural BAC or source changes need EV refresh. No Rebuild All or Snapshot EV.

## Worksheet protection

The workbook README explains the existing `okmd` password. It protects editing,
not confidentiality. Manual password changes in Excel can be reset by later
processing paths applying shared protection. Custom passwords are not a V1.1 feature.


## BU-1: update BOQ Amounts in a Mapping project

Engineering handoff; native Desktop/Excel and PO acceptance pending.

1. Open the existing saved Mapping project. In Excel, recalculate and Save the
   revised BOQ first, then select **Update BOQ Amounts...** in Mapping.
2. Select the revised workbook. The existing selected BOQ worksheet is compared
   against the verified project snapshot. Review changed items, affected Activity
   Amounts, mapped/unmapped counts, totals and formula usage before confirming.
3. **Apply and Save** preserves allocation pairs and percentages, saves a sibling
   `*.before-boq-update-*.progressstudio` recovery copy of the current project,
   then atomically saves the updated project. Cancel makes no project changes.
   An unchanged candidate cannot be applied. Keep the recovery copy until review
   is complete; open it to recover the pre-update project.
4. Save/close/reopen as usual and use existing Mapping export to produce the
   updated Progress Workbook. Existing exported Workbooks and XML are unchanged.

Only positive-to-positive Amount changes are supported. Row insertion, deletion,
reordering, changed descriptions/hierarchy/units/quantities or other non-Amount
cells, ambiguous duplicate identities and eligibility transitions are rejected.
Comparison covers every content row/column of the selected Mapping sheet; it is
not structural reconciliation of other worksheets. Formatting is not identity.
Non-Amount formulas and their saved values must remain unchanged.

Saved numeric formula results are supported and formula counts appear in review.
Missing, invalid or error results block the update. An unchanged heading with an
explicit saved empty-string formula result remains a non-item. Formula cache
freshness cannot be independently guaranteed; Excel recalculation/Save is required.

The updated BOQ is embedded in the project and pinned to reviewed file bytes.
If an external BOQ is changed or resaved, reopening uses the embedded reviewed
snapshot rather than trusting changed formula caches. Use Update BOQ Amounts
again for a later revision. Legacy projects need a verified original BOQ or an
embedded snapshot. Failed validation or persistence retains the live project;
a recovery copy may remain if the final write fails. A post-save display error
explicitly reports that the save succeeded and asks you to reopen the project.

## Export Activity Amounts to Primavera P6 XML

Use **Tools > Export Amounts to P6 XML...** without a Project Session or BOQ.
Select the saved Workbook explicitly. If amounts were changed in Mapping, first
Export Mapped Workbook and select that output. Choose the target XML and named
Activity Cost field (or reviewed creation). Confirm monetary intent, recalculate
formulas in Excel and Save when applicable, then Review Amounts and warnings.
Export XML opens Save for a new file; inputs and existing outputs are never
overwritten. Import the new file into P6 separately. See the
[writeback workflow](XML_AMOUNT_WRITEBACK.md#desktop-operation--xw-2) for review,
formula-cache limitations, cancellation and recovery. XW-2 native UX/PO acceptance
is pending; XW-1 compatibility acceptance remains accepted.
