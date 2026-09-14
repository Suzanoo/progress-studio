# FF-1 Finance Model contract

FF-0 and its Desktop Excel gate are Product Owner accepted. This milestone adds
only source-neutral reference calculations. No reader, renderer, workbook formula,
chart, GUI entry, calculation policy or existing service composition is changed.

## Architecture and reuse

`domain/financial_forecast.py` uses the repository's frozen dataclass pattern.
`services/financial_forecast_deriver.py` provides a pure `derive(FinanceInputs)`
service and small calculation helpers, with `FinanceValidationError(ValueError)`.
Existing Payment records describe progress requirements, not cash settlements;
existing EV uses Contract Value and cannot supply a cost budget. Neither is
repurposed. No existing Decimal finance utility was found in domain/services;
rounding is local to this module. There are no new dependencies or Excel imports.

## Monetary and time basis

- One declared project currency and cash-account boundary per input. The caller
  must supply consistent net cash/tax treatment and genuine contract amounts.
  V1 supports two decimal currency units only; no FX, tax engine or accruals.
- Amounts use Decimal, round half-up to 0.01 at entry. Decimal-compatible numeric
  strings/integers/floats are accepted (floats via decimal text); booleans,
  nonfinite values and individual magnitudes above 1e18 are rejected. Prefer
  Decimal at adapters. All forecast balances are nonnegative. Opening cash may
  be signed. Service/helper calculation contexts use precision 50.
- Dates are exact calendar dates, without times. Customer terms are calendar
  days, not months or business days. Finance actuals completeness is independent
  of EV_View_Date, Dashboard cutoff, and the progress phasing cutoff.
- Opening cash is immediately before events on opening_date. Cash movement
  totals and the balance series include actual events on/after that date through
  actuals_through inclusive. Earlier supplied history affects settlements and
  the paid-cash AC proxy, but is already embedded in opening cash.
- Actual records after actuals_through are reported as `future_actual:<id>` and
  excluded; they are not silently converted into forecast records. Moving a
  cutoff is recalculation using the supplied records, not a historical snapshot.

## Actual cash, outstanding balances and overrides

CashEvent has a unique stable ID, exact date, direction, amount, operating or
financing category, and optional forecast item link. Positive amounts move the
chosen stream. A negative amount reverses that stream and requires an explicit
correction reason. Corrections are manual adjustments, not an automatic reversal
or immutable accounting audit subsystem. Duplicate IDs and unknown/mismatched
settlement links fail. Same amount/date with different IDs is not automatically
identified as a duplicate; the adapter/user owns business identity reconciliation.

ForecastItem represents a disjoint obligation or remaining estimate:

| balance_date | Meaning of amount | Settlement deducted |
| --- | --- | --- |
| None | Original obligation | All linked actuals through cutoff |
| Explicit date | Remaining balance at END of that date | Linked actuals strictly after balance_date, through cutoff |

This distinction prevents a seeded remaining balance from deducting historical
payments twice. Future snapshots are rejected. Net deducted settlement must be
between zero and the supplied amount; overpayment or net negative settlement
requires reconciliation rather than being clamped. Partial settlement leaves
only unpaid cash. Retention items are operating receivables and may be seeded
using the same end-of-day snapshot convention.

An override applies AFTER settlement and replaces the item's remaining amount
and/or expected cash date. None keeps the base; zero cancels. `clear_date=True`
explicitly removes a date. A nonblank reason is mandatory. Result records retain
original item amount/date, net settlement, base remaining, effective remaining,
effective date and reason. The difference between base/effective amount is the
visible scenario adjustment; overrides never create a second item or edit cash.

Positive outstanding cash without a date is `undated`; a due date on/before the
actuals cutoff is `overdue`. Both remain in remaining totals and result records,
but are excluded from future dated balances. Nothing is rolled forward silently.
A zero remaining item is `settled` and does not generate a missing-date warning.

## Receipt and payment forecast helpers

BOQ Amount basis is Contract Value. No CPI, CV, EAC, Activity P/L, or cash-to-Activity
allocation is introduced. Cash Out AC proxy is the net OPERATING Cash Out in all
supplied history through cutoff, including explicit corrections. Financing is
separate and excluded from this proxy. Incomplete history makes this a partial
proxy; opening cash cannot reconstruct missing historic cost. Terminal cash is
not profit.

`remaining_unperformed_value(contract_value, certified_gross, earned_uncertified)`
returns Contract Value minus the two disjoint gross buckets, rejecting a negative
result. Certified gross includes settled receipts and retention; known remaining
certified receivables, earned-but-uncertified work, and future unperformed work
must be supplied separately. The helper does not infer certification from EV.

`certificate_items(id, gross, certification_date, ...)` creates `<id>:net` and
`<id>:retention`. Default credit_days=30 and retention_rate=0.05. Withheld cash
is rounded once; net=gross-withheld, conserving gross exactly. Jan 31, 2026 + 30
days is March 2, 2026. Retention release date must be explicit or stays undated;
project finish is never used as a release trigger. These defaults apply only to
eligible gross modeled certificates, never already-net manual receipts or
supplier payments. Opening retention + new withholding - linked release is
represented by separate retention items/events with no second withholding.

`remaining_cash_out(total_outturn, paid_to_date, known_outstanding)` returns the
uncommitted residual T-O-K. Negative residual is rejected, not floored. T is an
explicit cash estimate, not BAC/CPI or accounting EAC. Known commitments remain
separate from the residual. Missing estimates require cash_out_complete=False.

`phase_remaining(amount, weighted_dates, progress_through=...)` allocates a
separately reconciled remaining amount across unique, explicitly supplied future
dates with nonnegative weights. Largest-remainder cents conserve the amount;
earliest date breaks ties. Positive remainder without positive future weights
fails and requires revised timing. This is an allocation scenario, not delay
prediction. For modeled receipts callers may supply calendar month-end
certification dates, then call certificate_items to apply terms. The helper does
not apply customer terms to outflows. Schedule/BOQ extraction and readiness
validation belong to future adapters, not this calculation service.

## Balance, funding and monthly reporting

Balance = opening cash + cumulative actual/forecast inflows - outflows. Dated
points aggregate all streams on the same calendar day, then evaluate the closing
balance. The minimum also considers opening cash. Peak funding requirement =
max(0, -minimum_balance), using a zero buffer. Earliest occurrence wins a minimum
tie; an opening minimum is reported with opening_date. Intraday order is unknown.
This is the incremental deficit of the supplied scenario: actual and explicitly
forecast financing affect cash, but forecast financing remains in forecast
columns and is never treated as received. Availability of future financing is
not verified by the model.

Monthly points sum incremental actual/forecast streams by calendar month and
carry forward closing cash through empty months. Partial opening months include
only cash from opening_date. Monthly aggregation does not determine peak funding.
Reporting extends through the latest dated cash, cutoff, or horizon_end, whichever
is latest. A horizon cannot clip a late receipt or retention release.

Result operating remaining totals include scheduled, overdue and undated items;
retention is a subset of remaining receivables. Financing outstanding remains in
the resolved item records and cash series, separately categorized. Each resolved
item exposes its direction/category/status so adapters can aggregate undated or
financing balances without inventing workbook-dependent logic.

## Completeness and limits

forecast_complete requires explicit actuals_complete, receivables_complete and
cash_out_complete declarations plus no unresolved item/event issues. Unlinked
actual cash requires explicit unlinked_actuals_reconciled confirmation. The model
cannot prove that omitted transactions, invoices or commitments do not exist.
These declarations mean the caller has reconciled disjoint balances and source
basis; they are not inferred from a last nonzero cell. Undated/overdue balances
and future actual records make dated funding conditional/incomplete even if the
three declarations are true.

FF-1 tests establish arithmetic, date boundaries, validation and conservation.
Excel/Python formula parity, input persistence, live F9 recalculation, performance
on large workbook inputs, workflow integration and presentation remain unproven
by FF-1 and belong to later authorized milestones. FF-0 acceptance is retained.
