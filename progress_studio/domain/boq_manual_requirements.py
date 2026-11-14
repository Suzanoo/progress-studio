"""Manual planned BOQ requirements: disjoint composition over unchanged BP-1."""
from dataclasses import dataclass, replace
from decimal import Decimal
from progress_studio.domain.boq_payment_projection import Projection, ItemPeriod, ItemProjection, Diagnostic

D = Decimal


@dataclass(frozen=True)
class ManualInput:
    key: str
    values: tuple[D | None, ...]
    source_row: int
    first_column: int = 9


@dataclass(frozen=True)
class ManualPeriod(ItemPeriod):
    raw_period: D | None = None
    source_row: int = 0
    source_column: int = 0
    state: str = 'Unknown'


@dataclass(frozen=True)
class ManualItem(ItemProjection):
    origin: str = 'Manual BOQ'


@dataclass(frozen=True)
class CompositeProjection(Projection):
    mapped_projection: Projection | None = None


def manual_periods(entry, record, period_ids):
    if len(record.values) != len(period_ids):
        raise ValueError('Manual period count mismatch: ' + record.key)
    known_prefix, cumulative = True, D(0)
    result = []
    amount = entry.item.amount
    money_valid = amount is not None and amount.is_finite() and amount >= 0
    for index, (period_id, raw) in enumerate(zip(period_ids, record.values)):
        if raw is not None and (not isinstance(raw, D) or not raw.is_finite() or not -1 <= raw <= 1):
            raise ValueError(f'Invalid signed Period %: {record.key} / {period_id}')
        known_prefix = known_prefix and raw is not None
        if known_prefix:
            cumulative += raw
            if not 0 <= cumulative <= 1:
                raise ValueError(f'Manual cumulative outside 0..100%: {record.key} / {period_id}')
        current = cumulative if known_prefix else None
        state = 'Explicit' if known_prefix else 'Unknown' if raw is None else 'Incomplete history'
        warnings = ('Negative manual correction',) if raw is not None and raw < 0 else ()
        result.append(ManualPeriod(
            period_id, D(0), D(1) if raw is not None else D(0), current, current, raw,
            amount * current if money_valid and current is not None else None,
            amount * current if money_valid and current is not None else None,
            amount * raw if money_valid and raw is not None else None,
            'Manual BOQ; Unmapped' + ('; Incomplete history' if not known_prefix else ''),
            warnings, (), raw, record.source_row, record.first_column + index, state))
    return tuple(result)


def compose(projection, records):
    """Never invent shares/contributors or modify the accepted mapped projection."""
    if records is None:
        return projection
    by_key = {r.key: r for r in records}
    if len(by_key) != len(records):
        raise ValueError('Duplicate manual BOQ identity')
    catalog = {entry.item.key: entry for entry in projection.items if entry.item.role == 'item'}
    for key in by_key:
        if key not in catalog:
            raise ValueError('Manual BOQ identity missing from catalog: ' + key)
        if any(p.contributions or p.mapping_coverage != 0 for p in catalog[key].periods):
            raise ValueError('Manual/Mapping conflict: ' + key)
    entries = tuple(ManualItem(e.item, manual_periods(e, by_key[e.item.key], projection.periods))
                    if e.item.key in by_key else e for e in projection.items)
    diagnostics = []
    manual_entries = [e for e in entries if getattr(e, 'origin', '') == 'Manual BOQ']
    valid = [e for e in manual_entries if e.item.amount is not None and e.item.amount.is_finite() and e.item.amount >= 0]
    total = sum((e.item.amount for e in valid), D(0))
    for index, period_id in enumerate(projection.periods):
        period_known = [e for e in valid if e.periods[index].raw_period is not None]
        cumulative_known = [e for e in valid if e.periods[index].cumulative_fraction is not None]
        diagnostics.append(Diagnostic('ManualCoverage', period_id,
            f'Manual origin: {len(period_known)}/{len(manual_entries)} rows with known Period %; '
            f'{len(cumulative_known)}/{len(manual_entries)} with known cumulative prefix; '
            f'valid Amount coverage period={sum((e.item.amount for e in period_known), D(0))}/{total}, '
            f'cumulative={sum((e.item.amount for e in cumulative_known), D(0))}/{total}; '
            f'{len(manual_entries)-len(valid)} invalid Amount rows; no Activity Mapping implied'))
    return CompositeProjection(projection.policy, projection.periods, entries,
        projection.diagnostics + tuple(diagnostics) + (Diagnostic('ManualBOQ', 'Planned requirements',
            f'{len(records)} wholly unmapped input rows; blank is unknown; signed Period % of full BOQ Amount'),),
        projection.provenance, projection.period_sources, projection.activity_names, projection)
