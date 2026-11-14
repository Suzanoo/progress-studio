"""Read-only BOQ Requirement projection. No eligibility or accounting semantics."""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from datetime import date
from typing import Any

D = Decimal
TOLERANCE = D('1e-11')  # MappingStore 1e-9 percentage points, in fraction units.


def number(value):
    if value is None or isinstance(value, bool):
        raise ValueError('Missing or invalid number')
    try:
        result = D(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError('Invalid number') from exc
    if not result.is_finite():
        raise ValueError('Non-finite number')
    return result


@dataclass(frozen=True)
class CatalogRow:
    key: str
    stable_id: str
    source_sheet: str
    source_row: int
    project_row: int
    wbs: str
    hierarchy: tuple[str, ...]
    description: str
    unit: Any
    quantity: Any
    raw_amount: Any
    amount: Decimal | None
    amount_status: str
    role: str = 'item'


@dataclass(frozen=True)
class Diagnostic:
    code: str
    identity: str
    detail: str


@dataclass(frozen=True)
class Contribution:
    activity_id: str
    share: Decimal
    raw_requirement: Decimal | None
    derived_requirement: Decimal | None
    state: str
    source_row: int | None
    source_column: int | None
    weighted_fraction: Decimal | None


@dataclass(frozen=True)
class ItemPeriod:
    period_id: str
    mapping_coverage: Decimal
    input_coverage: Decimal
    known_fraction: Decimal | None
    cumulative_fraction: Decimal | None
    incremental_fraction: Decimal | None
    known_amount: Decimal | None
    cumulative_amount: Decimal | None
    incremental_amount: Decimal | None
    status: str
    warnings: tuple[str, ...]
    contributions: tuple[Contribution, ...]


@dataclass(frozen=True)
class ItemProjection:
    item: CatalogRow
    periods: tuple[ItemPeriod, ...]


@dataclass(frozen=True)
class PeriodSource:
    period_id: str
    source_column: int
    payment_date: date | None


@dataclass(frozen=True)
class Projection:
    policy: str
    periods: tuple[str, ...]
    items: tuple[ItemProjection, ...]
    diagnostics: tuple[Diagnostic, ...]
    provenance: tuple[tuple[str, str], ...] = ()
    period_sources: tuple[PeriodSource, ...] = ()
    activity_names: tuple[tuple[str, str], ...] = ()


def project(catalog, allocations, payment, activity_ids):
    """Pure deterministic projection; malformed identities/shares fail closed.

    Money is never rounded here. An incomplete item cannot produce a final %,
    even if its known contribution is numeric zero. Hierarchy rows are retained
    as catalog context and never counted as additional monetary items.
    """
    ids = set(activity_ids)
    if len(ids) != len(activity_ids):
        raise ValueError('Duplicate Activity identity')
    by_key = {x.key: x for x in catalog}
    stable = [x.stable_id for x in catalog if x.stable_id]
    if len(by_key) != len(catalog) or len(stable) != len(set(stable)):
        raise ValueError('Duplicate BOQ identity')
    periods = tuple(x.period_id for x in payment.periods)
    if not periods or len(set(periods)) != len(periods):
        raise ValueError('Missing or duplicate periods')
    if len(set(payment.activity_ids)) != len(payment.activity_ids) or not set(payment.activity_ids) <= ids:
        raise ValueError('Invalid Payment Activity identities')
    pairs = {}; shares = {}
    for a in allocations:
        if a.boq_key not in by_key or a.activity_id not in ids:
            raise ValueError('Mapping references missing BOQ or Activity')
        if by_key[a.boq_key].role != 'item':
            raise ValueError('Mapping references hierarchy rather than item')
        share = number(a.share_percent) / 100
        if not 0 < share <= 1 or (a.boq_key, a.activity_id) in pairs:
            raise ValueError('Invalid share or duplicate mapping pair')
        pairs[a.boq_key, a.activity_id] = share
        shares[a.boq_key] = shares.get(a.boq_key, D(0)) + share
        if shares[a.boq_key] > 1 + TOLERANCE:
            raise ValueError('Overallocated BOQ item')
    entered = {}
    for period in payment.periods:
        for r in period.requirements:
            key = (r.activity_id, period.period_id)
            value = number(r.required_fraction)
            if r.activity_id not in ids or key in entered or not 0 <= value <= 1:
                raise ValueError('Invalid or duplicate Requirement')
            entered[key] = (value, r.source_row, r.source_column)
    states = {}; diagnostics = []
    for a in sorted(ids):
        known = any((a, p) in entered for p in periods)
        latest = D(0); last_row = last_col = None; started = False; warning = False
        for p in periods:
            raw = entered.get((a, p))
            if raw:
                if raw[0] < latest:
                    warning = True
                    diagnostics.append(Diagnostic('DecreasingRequirement', a, p))
                latest, last_row, last_col = raw; started = True
            state = 'Explicit' if raw else 'Carried' if started else 'BeforeFirst' if known else 'Unknown'
            states[a, p] = (latest if known else None, state, last_row, last_col, raw[0] if raw else None, warning)
    result = []
    for item in catalog:
        if item.role != 'item':
            result.append(ItemProjection(item, ())); continue
        mapping = shares.get(item.key, D(0)); links = sorted((a,s) for (b,a),s in pairs.items() if b == item.key)
        if mapping == 0:
            diagnostics.append(Diagnostic('Unmapped', item.key, f'BOQ Amount={item.amount}'))
        elif abs(mapping-1) > TOLERANCE:
            diagnostics.append(Diagnostic('PartialMapping', item.key, f'Coverage={mapping}'))
        elif mapping != 1:
            diagnostics.append(Diagnostic('ShareToleranceResidual', item.key, f'Unscaled coverage={mapping}'))
        money = item.amount
        if money is not None:
            money = number(money)
            if money < 0: money = None
        if money is None:
            diagnostics.append(Diagnostic('AmountUnavailable', item.key, item.amount_status))
        rows = []; previous = D(0)
        for p in periods:
            trace = []; coverage = D(0); k = D(0); warnings = []
            for a,share in links:
                value,state,row,col,raw,decreasing = states[a,p]
                contribution = share * value if value is not None else None
                if value is not None: coverage += share; k += contribution
                if decreasing: warnings.append('DecreasingRequirement')
                trace.append(Contribution(a,share,raw,value,state,row,col,contribution))
            complete = abs(mapping-1) <= TOLERANCE and coverage == mapping
            final = k if complete else None
            delta = final-previous if final is not None and previous is not None else None
            known = k if coverage > 0 else None
            status = 'Unmapped' if mapping == 0 else 'PartialMapping' if abs(mapping-1)>TOLERANCE else 'IncompleteRequirement' if not complete else 'Complete'
            mult = lambda x: money*x if money is not None and x is not None else None
            rows.append(ItemPeriod(p,mapping,coverage,known,final,delta,mult(known),mult(final),mult(delta),status,tuple(sorted(set(warnings))),tuple(trace)))
            previous = final
        result.append(ItemProjection(item,tuple(rows)))
    return Projection('bp1-derived-requirements-v1',periods,tuple(result),tuple(diagnostics),
                      period_sources=tuple(PeriodSource(p.period_id,p.column_index,p.payment_date) for p in payment.periods))

