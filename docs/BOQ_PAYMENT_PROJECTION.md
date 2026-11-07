# BP-1 BOQ Payment Projection

Status: BP-1 accepted; stable tag `bp1-boq-payment-projection-stable`.
Baseline: `33512821d059d99fbf2c90dda207a6c4d2127dd4`.

## Scope and entry point

BP-1 provides a read-only Python service. It does not add a desktop action,
write an Excel report, change Payment Input, or calculate certified payment.
BP-2 consumes this service without changing projection semantics. See
[BOQ Payment Export](BOQ_PAYMENT_EXPORT.md) for the separate snapshot workbook.

```python
from pathlib import Path
from progress_studio.services.boq_payment_projection import BOQPaymentProjectionService

result = BOQPaymentProjectionService().calculate(
    Path('project.progressstudio'),
    Path('current_payment.xlsx'),
    accept_revision=True,  # explicitly select current revision; diagnostics remain
    boq_path=Path('original_boq.xlsx'),  # optional; otherwise use saved snapshot
)
```

Read `result.diagnostics` and `result.provenance` alongside `result.items`.
`accept_revision` defaults to false. It is an acknowledgement of selecting a
changed Payment workbook, not permission to relink a Mapping Session. The
existing relinking validator is unchanged. No new session persistence is added.

## Sources and compatibility

Saved explicit allocation percentages are authoritative, expressed as 0–100
percentage points. Activity Amount never creates or changes an allocation.
The service uses existing session loading, verified embedded source snapshots,
working-tree restoration, BOQ identity helpers and the sparse Payment reader.

Both snapshots must agree with their saved source fingerprints. An optional
selected BOQ must pass the existing `validate_workbook` contract. Mapped Activity
IDs must exist in both the saved working tree and current main. Current main
names must match saved source main names; edited Mapping labels are diagnosed
separately. Duplicate IDs, absent mapped identities, overallocations, malformed
shares and invalid Requirements fail closed. Date/WBS revisions are visible.

This first service supports saved sessions with embedded source snapshots and
without supplemental Activity/WBS records. Older snapshot-less sessions and
supplemental sessions fail explicitly; it never guesses or asks for remapping.
Their existing Mapping/relink workflows remain unchanged. Matching IDs/names
cannot prove real-world project identity if a different project deliberately
reuses all identifiers; source selection and hash provenance remain essential.

Input file hashes are checked before and after calculation. No source workbook
or session is saved. Temporary snapshot copies are removed after use.

## Calculation contract

For each Activity, use explicit saved Requirements when present. If any valid
Requirement exists in its schedule, before-first periods derive zero and later
blanks carry the last entered value. A schedule with no entries is unknown.
An explicit zero is an entry. Input blanks are never filled or persisted.

For BOQ item j and period p, known fraction K is the sum of share(j,a) times
derived Requirement(a,p) over known mapped Activities. Mapping coverage is the
sum of explicit shares; input coverage is the sum of shares with known states.
Neither uses Activity Amount as a weight. Shares are never normalized.

A final cumulative fraction is available only when the item is fully mapped
and all mapped Requirements are known. Partial and unmapped items retain their
known contributions/coverage but have no final fraction or final Amount.
Incremental fraction is current minus previous cumulative fraction, using zero
as the opening baseline. Negative changes are preserved with decrease warnings.
Unknown results are `None`, not numeric zero.

Money equals BOQ Amount times the corresponding fraction. Missing, invalid,
negative and unresolved Amounts make money unavailable while valid percentages
remain inspectable. Zero Amount yields zero money. Unmapped Amount and Activity
Amount reconciliation are separate diagnostic families, never a combined balance.

Decimal arithmetic retains source precision without cent-rounding per Activity.
Mapping completeness uses the existing tolerance of 1e-9 percentage points;
tolerated residuals are reported and never rescaled. Downstream display rounding
must not be fed back into calculations. Decimal finite-precision conservation
is tested with tolerances far below a monetary cent.

## Result and traceability

Immutable result records contain policy version, period IDs, period dates/source
columns, catalog identity/source coordinates, BOQ hierarchy, amount validity,
status, coverage, known/final cumulative and final incremental values. Every
item-period includes Activity ID, share, raw/derived Requirement, state, source
cell and weighted contribution. Carried states reference their last entered cell;
before-first derived zero has no entered source cell.

Hierarchy rows are retained as context with no item-period monetary results.
They are not additional BOQ items, so summing leaf items avoids double counting.
A blank Amount/Unit/Qty hierarchy label (including supported blank source links)
is retained as a group. Arbitrary unresolved formulas are not evaluated.
Cached formula values are snapshots and cannot prove Excel cache freshness;
recalculate and save source files before selecting a revision when necessary.

## Verification

```text
python -m pytest -q tests/unit/payment/test_boq_payment_projection.py tests/integration/payment/test_boq_payment_projection_service.py
python -m pytest -q
python scripts/bp1_projection_proof.py "NKC1-R02(2).progressstudio" "NKC1_R2_payment(1).xlsx" "BOQ  NKC1-BOQStudio-Output(1).xlsx" --output NKC1_PROOF.json
```

The proof script is deliberately specific to the supplied NKC1 fixture layout;
it is an independent raw-cell oracle, not a general workbook export interface.
It checks all mapped item-period results, signed differences, conservation,
traceable example and unchanged source hashes. No native Excel rendering is
introduced by BP-1. Existing PI-1, Payment Position, Snapshot/Live Rebuild,
Mapping, EV, BOQ Update and XML behavior remain owned by existing services.
