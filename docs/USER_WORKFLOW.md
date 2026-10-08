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
For Amount, choose **Select Amount field / Preview...**, select the intended
numeric custom field, review its raw values/counts/total, and confirm the selection.
No field is chosen automatically. No fallback amount needs to be entered.

Equal gives each ordinary activity one calculation unit. Duration uses the
source working duration normalized to hours (P6 PlannedDuration / MSP Duration),
not elapsed Start–Finish time. Milestones stay in the workbook with zero weight;
summary/WBS rows roll up activity weights. Invalid ordinary duration stops
Duration creation with the affected Activity IDs; it never switches to Equal.

Equal/Duration dummy weights are not Contract Value, actual cost or cash flow.
Amount uses real monetary values from the selected XML field; this does not
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

## 6. Recalculation in Excel

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
