# Progress Studio Architecture

This file is the technical source of truth for current Progress Studio
ownership boundaries. Historical milestone documents under
`docs/history/` describe how the product reached this architecture but
do not override this contract.

## 1. End-to-end architecture

``` mermaid
flowchart TD
    MSP[MS Project XML] --> DETECT[XML Format Detector]
    P6[Primavera P6 XML] --> DETECT

    DETECT --> ADAPTER[MSP / P6 Adapter]
    ADAPTER --> NORMAL[Normalized Schedule]
    NORMAL --> VALIDATE[Validation]
    VALIDATE --> CREATE[Create Progress]

    CREATE --> MAIN[main\nWorkbook Source of Truth]
    CREATE --> MONTHLY[main_monthly]
    CREATE --> DASH[Dashboard]
    MAIN --> MAP[Mapping]
    MAIN --> PAYMENT[Payment Workspace]
    MAIN --> EDIT[User edits Excel]

    PAYMENT --> STDPAY[Standard Payment]
    PAYMENT --> PBD[Payment Breakdown]

    EDIT --> REBUILD[Rebuild Workspace]
    REBUILD --> RPROG[Progress Rebuild]
    REBUILD --> RPAY[Payment Rebuild]

    RPROG --> RENDER[Owned Renderers]
    RPAY --> RENDER
    MAP --> FINAL[Final Workbook Policy]
    STDPAY --> FINAL
    PBD --> FINAL
    RENDER --> FINAL

    FINAL --> SAVE[Save Workbook]
```

## 2. Source boundaries

### Before workbook creation

Schedule XML is normalized through source-specific adapters:

``` text
MSP XML ─┐
         ├─> source adapter -> Normalized Schedule -> validation
P6 XML ──┘
```

The Progress workbook engine should not need to know which XML dialect
produced the normalized schedule.

Supported XML paths in the current normalizer:

-   Microsoft Project XML.
-   Primavera P6 XML.

Amount/cost is deliberately not normalized from schedule XML. Initial
workbook generation keeps the established fallback/fake amount behavior;
BOQ Mapping owns real cost allocation.

### After workbook creation

`main` becomes the workbook source of truth for the editable
schedule/progress state.

Rebuild operates from the selected workbook and does not require the
original XML, BOQ file, mapping session, or GUI tree as an input to the
standalone rebuild contract.

Payment Breakdown also derives from current `main` and does not require
the original XML, BOQ file, Mapping session, or `Payment Input`.

## 3. Ownership matrix

  ------------------------------------------------------------------------
  Component               Owns                     Must not own
  ----------------------- ------------------------ -----------------------
  XML Detector / Adapters XML dialect detection    Excel rendering,
                          and source normalization Mapping, Payment

  Normalized Schedule     Source-neutral           Workbook formatting
                          WBS/Activity schedule    
                          contract                 

  Create Progress         Initial workbook         BOQ allocation rules,
                          generation               post-edit rebuild
                                                   decisions

  Mapping                 BOQ -\> Activity         Payment rendering,
                          allocation / Activity    Progress rebuild
                          Amount                   

  Standard Payment        Payment Input            Progress/Dashboard
                          reconciliation and       regeneration, Payment
                          `Payment` rendering      Breakdown grouping

  Payment Breakdown       exact-name grouping from Payment Input
                          current `main` and       reconciliation,
                          `Payment-Breakdown`      standard Payment
                          rendering                rendering, Progress
                                                   rebuild

  Progress Rebuild        Progress-derived sheets  Payment Input ownership
                          from current `main`      

  Payment Rebuild         `Payment` output from    Progress-derived
                          current                  sheets,
                          `main + Payment Input`   `Payment-Breakdown`

  Renderers               Presentation objects     Other renderer
                          they create              ownership / business
                                                   calculation

  Final Workbook Policy   guide, visibility,       Progress/Payment
                          protection and Excel     calculation
                          recalc policy            

  Excel / user            editable workbook inputs Python snapshot
                          and user-created sheets  generation
  ------------------------------------------------------------------------

## 4. Create Progress contract

The initial pipeline builds the workbook once from normalized schedule
data and then applies presentation/final workbook policy before saving.

Conceptually:

``` text
Normalize
  -> schedule/activity workbook data
  -> weekly timescale
  -> progress/distribution
  -> monthly view
  -> dashboard / overlays
  -> final workbook policy
  -> save
```

Create Progress owns the initial reporting labels and display margins.

## 5. Timescale contract

Display range and reporting range are different concepts.

``` text
DISPLAY RANGE
X  X  X | W1 W2 W3 ... Wn | X X X
          REPORTING RANGE

MONTHLY
X | M1 M2 M3 ... Mn | X
```

-   `X` is a display-only margin period.
-   `W1` is the first weekly period overlapping the project reporting
    window.
-   `Wn` is the final weekly period overlapping the project finish.
-   `M1...Mn` follow the same rule at monthly level.
-   The physical date is authoritative for calculations.
-   W/M labels are human-facing reporting metadata, not calculation
    identity.

The display margin may remain visible in `main` / `main_monthly`. It
must not extend Dashboard reporting data or progress calculation ranges.

## 6. Reporting-range contract

Reporting-derived outputs use only periods that overlap the project
schedule window.

``` text
Display:    X X | project reporting periods | X X
Reporting:      | project reporting periods |
```

This applies to Dashboard reporting sources, chart series boundaries,
cumulative Plan/Actual reporting and other derived reporting views.

The final overlapping reporting period is retained even when Project
Finish occurs inside that period.

## 7. Rebuild workspace contract

Rebuild is an orchestrator with a 2 x 2 user-facing matrix:

  -----------------------------------------------------------------------
  Workbook mode           Progress                Payment
  ----------------------- ----------------------- -----------------------
  Snapshot                rebuild Progress-owned  rebuild standard
                          generated outputs       Payment only

  Live                    rebuild Live            rebuild Live standard
                          Progress-owned outputs  Payment only
  -----------------------------------------------------------------------

### Progress-owned outputs

Progress rebuild may replace the generated Progress views required by
the selected mode. `main` remains authoritative and is preserved as the
edited source.

### Payment-owned output

Payment rebuild owns `Payment` only and uses current
`main + Payment Input`. Progress-derived views are preserved.

`Payment-Breakdown` is intentionally outside the Rebuild ownership
matrix. It is generated explicitly from the Payment Workspace using
current `main`.

### W/M labels during rebuild

Rebuild does **not** renumber weekly `W/X` labels in `main`. User
workbook structure is preserved. Rebuild readers/calculations should use
dates and physical structure rather than treating a W/M sequence number
as business identity.

When a monthly view is regenerated, it is rebuilt as monthly buckets and
follows the `X/M` display/reporting contract.

### Earned Value live-workbook boundary

EV Rebuild owns structural BAC / BOQ / mapping topology. After a successful
EV build, `main` remains the live authority for Plan and Actual progress.

Earned Value uses one semantic workbook control, `EV_View_Date`, for the
selected reporting view. PV and EV both accumulate live `main` values through
that view date. Dashboard cutoff controls are presentation state for Dashboard
and are not Earned Value calculation inputs.

Plan/Actual edits require only Excel recalculation (F9 / Save). Structural BOQ
or mapping changes require EV Rebuild. See `EV_LIVE_CONTRACT.md` for the full
contract.

## 8. Payment Breakdown contract

Payment Breakdown is a separate Payment Workspace path.

Its grouping identity is the complete Activity Name after
leading/trailing whitespace trim. Different full names remain different
groups.

It must not automatically use fuzzy, contains, keyword, case-normalized,
or Activity-ID matching.

Eligible source Activity progress profiles are preserved individually
and combined by Amount weighting:

\[ P_t = `\frac{\sum_i Amount_i \cdot P_{i,t}}{\sum_i Amount_i}`{=tex}
\]

Combined cumulative progress is derived from the combined period
profile.

This exact-name identity is an explicit feature exception. It does not
change the Activity-ID identity used by Mapping or standard Payment.

See `docs/PAYMENT_BREAKDOWN.md` for the complete feature contract.

## 9. Renderer ownership

A renderer owns only the workbook objects it creates.

Examples:

-   Traditional overlay renderer owns overlay series, cutoff redline,
    marker/label styling and overlay geometry.
-   Standard Payment renderer owns Payment lines/badges and its drawing
    objects.
-   Payment Breakdown renderer owns only the `Payment-Breakdown`
    worksheet it creates/replaces.
-   Dashboard renderer owns Dashboard chart/presentation objects.

A Payment workflow must not rebuild or restyle Progress overlays. Final
workbook policy may reassert portable workbook properties required to
survive openpyxl serialization, but it must not recalculate another
renderer's business data.

Standalone Payment Breakdown generation must preserve opaque workbook
package parts it does not own, including existing drawings and external
links.

## 10. Final Workbook Policy

All user-facing outputs converge on a shared final workbook policy for:

-   workbook guide / README sheet;
-   sheet visibility;
-   sheet protection and intended unlocked inputs;
-   Excel calculation properties.

F9 and Save belong to Excel formula recalculation. They do not execute
Python rebuild logic.

Python-owned generated snapshots/caches must be regenerated through the
owning Python workflow when required.

## 11. Performance policy

openpyxl workbook I/O is expensive. The design target is:

``` text
one user operation
  -> load/create workbook as few times as practical
  -> derive/mutate in RAM
  -> apply owned renderers/policy
  -> save once at the workflow boundary
```

Do not reopen a workbook solely to apply visibility, protection or
calculation properties when the active workbook object can be finalized
in memory.

Reopening a saved workbook for validation is appropriate in
tests/debugging, not as default production flow.

## 12. GUI and domain separation

-   GUI widgets are presentation state, never business data authority.
-   Mapping allocation belongs to domain/services, not Treeview rows.
-   Stable Activity identity is Activity ID unless a feature contract
    explicitly defines another identity.
-   Payment Breakdown is the explicit current exception: exact trimmed
    Activity Name is its grouping identity.
-   BOQ identity must use stable source metadata/keys rather than
    Description alone.
-   Unknown/user-created workbook sheets should be preserved unless an
    explicit contract says otherwise.

## 13. Architecture-change rule

A bug fix should not introduce a new architecture path when an existing
owner already exists.

Before changing cross-workspace behavior, tests should identify the
affected ownership boundary (Create, Mapping, Payment, Rebuild,
renderer, or Final Workbook Policy) and include the relevant workflow
regression gate.
