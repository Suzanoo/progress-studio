# Payment Breakdown

Payment Breakdown is an optional Payment Workspace output that derives
consolidated payment progress from the current `main` worksheet.

It is deliberately separate from standard Payment and from Rebuild.

## 1. Workspace and ownership

``` text
Payment Workspace
├─ Standard Payment
│  ├─ source: current main + Payment Input
│  ├─ identity: Activity ID
│  └─ output: Payment
│
└─ Payment Breakdown
   ├─ source: current main
   ├─ grouping identity: exact Activity Name
   └─ output: Payment-Breakdown
```

`Payment-Breakdown` is not a Rebuild-owned output. Payment Rebuild
continues to own `Payment` only.

## 2. Source of truth

Payment Breakdown is derived from the selected workbook's current `main`
data.

It does not require the original schedule XML, BOQ workbook, Mapping
session, or `Payment Input`.

Only eligible Plan Activity rows participate.

## 3. Exact Activity Name contract

Grouping identity is the complete Activity Name after trimming leading
and trailing whitespace.

Examples:

``` text
"First Fixed" == "First Fixed"
" First Fixed " == "First Fixed"

"U-Glass" != "U-Glass Wall"
"P3.1-P3.3 | U-Glass" != "P3.2 U-Glass Wall"
```

Payment Breakdown does not use:

-   fuzzy matching;
-   contains matching;
-   keyword regrouping;
-   case-normalized automatic matching;
-   Activity ID as the grouping identity.

If Activities should belong to one group but their names differ, correct
the names in `main` and build Payment Breakdown again.

This exception is local to Payment Breakdown. Standard Payment and
Mapping remain Activity-ID based.

## 4. Eligible source Activities

A source Activity is eligible when:

-   Activity ID is non-empty;
-   Activity Name is non-empty;
-   Amount is positive;
-   period progress values are non-negative;
-   the source progress profile totals 100%;
-   the exact trimmed Activity Name occurs at least twice for normal
    derived output.

A one-occurrence threshold is allowed only for diagnostic/prototype
review.

Invalid source rows are skipped rather than silently repaired.

## 5. Calculation contract

Each source Activity keeps its own period progress profile.

For Activities sharing one exact Activity Name, combined period progress
is Amount-weighted:

\[ P_t = `\frac{\sum_i Amount_i \cdot P_{i,t}}{\sum_i Amount_i}`{=tex}
\]

Combined cumulative progress is the cumulative sum of the combined
period progress.

The engine must not average Activity percentages equally when Activity
Amounts differ, and it must not mutate the individual source profiles
before weighting.

## 6. Output worksheet

The renderer creates or replaces only:

``` text
Payment-Breakdown
```

Each derived Activity group contains:

1.  group header;
2.  source Activity Progress rows;
3.  source Activity Cumulative rows;
4.  Combined Progress;
5.  Combined Cumulative.

Source Activity detail rows are Excel outline level 1 and are expanded
by default. The group header and Combined rows remain outside the detail
outline so the summary remains visible when details are collapsed.

Version 1 uses row outline only; it does not add column grouping.

## 7. Presentation contract

Presentation rules do not change the underlying calculated values.

### Progress rows

-   zero values display blank;
-   values between 0% and 100% display in red;
-   a genuine period value of 100% remains visible normally.

### Cumulative rows

-   leading zero values display blank;
-   values between 0% and 100% display in red;
-   the first 100% completion value remains visible;
-   repeated 100% values after completion display blank.

The same rules apply to source and Combined rows.

## 8. Workbook preservation

Payment Breakdown is an extension-only workflow and must preserve
workbook objects it does not own.

The standalone build path therefore:

1.  reads current `main`;
2.  derives the Payment Breakdown snapshot;
3.  works on a temporary copy;
4.  performs one mutable open/save boundary where practical;
5.  renders only `Payment-Breakdown`;
6.  reasserts required workbook portability/visibility properties;
7.  restores opaque OOXML drawing and external-link parts not owned by
    Payment Breakdown;
8.  validates the workbook package;
9.  atomically replaces the requested output.

This preservation rule is important because openpyxl serialization can
otherwise alter workbook package parts that belong to existing Progress,
Payment, Dashboard, or Earned Value outputs.

## 9. Relationship to Rebuild

Payment Breakdown is intentionally not part of the Rebuild 2 x 2 matrix.

``` text
Rebuild
                    Progress    Payment
Snapshot               ✓           ✓
Live                   ✓           ✓
```

Payment Rebuild continues to regenerate `Payment` from current
`main + Payment Input`.

To refresh `Payment-Breakdown` after editing Activity names, Amounts, or
progress in `main`, use **Build Payment Breakdown** in the Payment
Workspace.

Moving Payment Breakdown into Rebuild requires a separate architecture
decision because standard Payment owns drawing objects and has different
workbook-preservation requirements.

## 10. Engineering boundaries

Future changes must preserve these boundaries unless the architecture
contract is explicitly revised:

1.  do not replace exact-name grouping with Activity-ID, fuzzy,
    contains, keyword, or case-normalized matching;
2.  do not change standard Payment's Activity-ID identity contract;
3.  preserve each source Activity progress profile before Amount
    weighting;
4.  keep business calculation in services and worksheet presentation in
    the renderer;
5.  do not make `Payment-Breakdown` a Rebuild-owned generated sheet by
    accident;
6.  preserve workbook package parts not owned by Payment Breakdown;
7.  reuse the existing Payment Workspace and workbook-preservation
    patterns rather than introducing a parallel workflow.
