# FF-2 Finance inputs and live workbook view

Implementation baseline: main `8515608ccb983202ff7ee30b74a8036094db2438`.
The authoritative calculation definitions remain in FINANCIAL_FORECAST_CONTRACT.md.
FF-0 and FF-1 remain accepted. This delivery does not authorize FF-3.

## Run from the repository root

Create finance on a COPY of an existing Progress Studio workbook:

```text
python scripts/finance_workbook.py input.xlsx input_finance.xlsx --opening-date 2026-01-01 --actuals-through 2026-09-13 --currency THB
```

Choose the actual opening date, finance actuals-through date and currency for
your project. Opening cash initially defaults to zero; enter the reconciled
opening cash before using the result. Completeness declarations initially remain
FALSE. No cash, invoices, cost budget or transactions are inferred from Progress.

Refresh an existing finance workbook, retaining its inputs:

```text
python scripts/finance_workbook.py input_finance.xlsx refreshed_finance.xlsx
```

The output must be a new path. Source and existing output files are not overwritten.
Use `--capacity 200` to increase the prepared rows on refresh. Capacity is never
silently reduced. First creation defaults to 100 rows per input table. This is
a workbook/API/CLI milestone, not a new desktop workspace or Rebuild selector.

## Workbook ownership

- Finance Input: persistent settings, actual cash, forecast obligations and overrides.
- Cash Flow: live monthly cash and dated minimum/funding summary.
- Finance_Data: hidden, inspectable resolved items, dated cash, checks and calculations.

The three input tables start at row 25, side by side. Blue cells are editable.
Use the Notes columns for free-form notes. IDs and notes stay on the same record.
Do not insert/delete columns or edit headers. Prepared-row expansion is explicit.
The final visibility policy keeps Finance Input/Cash Flow visible and Finance_Data
normally hidden. Existing protection preserves the unlocked finance input cells.

Credit Days defaults to 30 calendar days; Retention Rate defaults to 5%.
Kind `certificate` treats Amount as original gross certification value and
Expected / certification date as certification date. It produces `ID:net` and
`ID:retention`, which are the IDs used for cash settlements and overrides.
Retention release must be explicit; project finish is never a release trigger.
Kind `manual` uses the entered net amount/date without applying customer terms.
Manual items can represent disjoint remaining receivables, cash-out estimates,
commitments, opening retention or explicitly forecast financing.

Balance date means the input amount is remaining at END of that date. Without
it the amount is an original obligation. Linked settlements are deducted using
FF-1's respective date rules. Override amount replaces the remaining amount
AFTER settlements. Blank keeps base; zero cancels. Date override and Clear date
are distinct. A reason is required. Overrides never rewrite actual cash.

Negative cash requires a correction reason. Financing is excluded from the
operating Cash Out AC proxy. BOQ Amount is Contract Value, not a cost budget.
No CPI/CV/EAC, activity P/L, cash-to-activity allocation, automated schedule
phasing or GUI feature is added. Users must reconcile the disjoint remaining
estimates before entering/confirming them; FF-1 helpers remain available to callers.

## Live calculation and limits

F9 and Save recalculate formulas; neither invokes Python. The source calculation
properties are preserved exactly by Finance extension/refresh. No manufactured
formula caches are written. First-open appearance is a separate Desktop Excel gate.

Funding uses opening cash and end-of-day balances on exact cash dates. Same-day
events are netted; the earliest minimum date wins. Monthly reporting aggregates
the same streams; month-end surplus cannot hide an earlier funding deficit.
Overdue/undated amounts remain in remaining totals but not future dated balances.
Future actual records are excluded and flagged, not converted to forecast.

Invalid inputs suppress headline financial outputs with #N/A and an INVALID INPUT
status. Incomplete declarations/unresolved cash show CONDITIONAL / INCOMPLETE.
Inspect Finance_Data's input-validation and item-status columns for details.

V1 input boundary: .xlsx, Excel 1900 date system, calendar dates from 1-Mar-1900,
amounts within +/-1e12 per entry, two-decimal currency. IDs use letters, digits,
underscore, dot, colon or hyphen, unique ignoring case. Excel validates wildcard
and whitespace ambiguities live; the reader enforces the full ID alphabet.
Inputs are values, not formulas. The prepared input limit is 2000 rows per table.
The monthly view has 600 prepared months; exceeding this is an explicit invalid
capacity condition, not a clipped successful forecast. Large-capacity Excel
performance and extreme aggregate monetary precision remain acceptance limits.

## Preservation and reuse

The implementation reuses FF-1 records/deriver, existing defined-name helper,
Dashboard font, final visibility/protection, and XLSX table validator. Mapping
uses the same edited-workbook precedence as Payment to copy finance inputs,
then rebuilds only finance-owned outputs. Other rebuild workflows retain them.

Finance extension requires a narrow package merge because the existing opaque
drawing-restoration helper identifies parts by number and cannot guarantee old
chart ownership after a general round-trip. Finance has no chart/drawing parts.
The new finance_package adapter grafts finance sheets/tables and their style/name
metadata into original bytes, leaving unrelated sheet/chart/drawing parts intact.
Existing style indexes are retained; finance styles are remapped/appended.
This is not a general workbook repair facility. Unsupported relationships on
finance sheets (such as inserted comments or hyperlinks) cause an explicit refusal;
use the ordinary Notes cells. Arbitrary third-party workbook features are unproven.

## Automated tests

```text
python -m pytest tests/unit/finance tests/regression/finance/test_dated_funding_contract.py tests/regression/finance/test_finance_lifecycle.py -q
python -m pytest -m smoke -q
python -m pytest -m regression -q
```

Optional independent formula evaluator, test environment only:

```text
python -m pip install formulas
python -m pytest tests/regression/finance/test_finance_formula_parity.py -q
```

The optional evaluator test inlines named references and resolves ISFORMULA from
cell metadata in a disposable copy because that evaluator does not implement
those features fully. Financial and validation arithmetic is unchanged. Without
the optional dependency this module is explicitly skipped, not passed.

## Desktop Excel acceptance

1. Open the supplied synthetic FF2_ACCEPTANCE.xlsx. No repair dialog is acceptable.
   Record first-open separately, then F9/Save. Expected: peak funding 50 on
   5-Feb-2026; minimum cash -50; remaining receivable 180; remaining cash out 100;
   retention 5. Monthly closing cash Jan-Jun: 50, 70, 125, 125, 125, 130.
2. Finance Input: U25 = cert:net, V25 = 0, Y25 = Cancel remaining. F9 must set
   remaining receivable to 125 and final cash to 75. Clear V25: receivable returns
   to 180 and final cash to 130. Actual receipt remains 40 in both cases.
3. Set B8 from 30 to 31: Finance_Data's cert:net date changes from 2-Mar to
   3-Mar-2026. Restore 30. Change B9 from 5% to 10%: net remaining becomes 50,
   retention 10; total remaining receivable remains 180. Restore 5%.
4. Clear the override, then enter an invalid link or negative cash without a
   correction reason. F9 must show INVALID INPUT and suppress headline figures.
   Restore valid values. Test an undated or overdue positive item: conditional
   status, amount still in remaining totals, no silent future date assignment.
5. Enter a zero override/reason and note on the last prepared row. Save/close/reopen,
   refresh to a new file, then run Progress Snapshot/Live, Payment Snapshot/Live,
   Mapping export and EV refresh on copies. Verify inputs, notes, zero overrides,
   editable cells and finance visibility persist. Compare existing views separately.
6. Repeat relevant checks with the real project workbook and its true opening cash,
   transactions and remaining estimates. Record Excel version, F9 timing and any
   first-open/chart/protection issues. Automated tests do not close this gate.
