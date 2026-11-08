"""BP-2 display aggregation over immutable BP-1 results; no eligibility logic."""
from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class ReportRow:
    key: str
    wbs: str
    name: str
    unit: object
    quantity: object
    amount: object
    status: str
    level: int
    group: bool
    values: tuple[tuple[object, ...], ...]


def report_rows(projection):
    items = projection.items
    codes = [x.item.wbs for x in items if x.item.wbs]
    if len(codes) != len(set(codes)):
        raise ValueError('Ambiguous duplicate BOQ WBS codes')
    leaves = [x for x in items if x.item.role == 'item']
    if not leaves:
        raise ValueError('No BOQ leaf items')
    fields = ('cumulative_fraction', 'incremental_fraction', 'cumulative_amount', 'incremental_amount')

    def summary(key, wbs, name, members, level):
        valid = all(x.item.amount is not None and x.item.amount >= 0 for x in members)
        total = sum((x.item.amount for x in members), Decimal(0)) if valid else None
        positive = [x for x in members if x.item.amount is not None and x.item.amount > 0]
        values = [[], [], [], []]
        complete = bool(members) and valid and bool(positive)
        for p in range(len(projection.periods)):
            ok = complete and all(x.periods[p].cumulative_amount is not None for x in positive)
            cumulative = sum((x.periods[p].cumulative_amount for x in positive), Decimal(0)) if ok else None
            delta_ok = ok and all(x.periods[p].incremental_amount is not None for x in positive)
            delta = sum((x.periods[p].incremental_amount for x in positive), Decimal(0)) if delta_ok else None
            for target, value in zip(values, (cumulative / total if ok else None,
                    delta / total if delta_ok else None, cumulative, delta)):
                target.append(value)
        status = 'Invalid Amount' if not valid else 'No positive BOQ Amount' if not positive else 'Incomplete' if any(v is None for v in values[0]) else 'Full'
        if any(p.warnings for x in members for p in x.periods):
            status += '; Decreasing Requirement'
        return ReportRow(key,wbs,name,'','',total,status,level,True,tuple(map(tuple,values)))

    # Reuse a genuine overall source group only if it contains every leaf.
    roots = [x for x in items if x.item.role != 'item' and x.item.wbs and
             all(y.item.wbs.startswith(x.item.wbs + '.') for y in leaves)]
    if len(roots) > 1:
        raise ValueError('Ambiguous overall BOQ hierarchy')
    result = [] if roots else [summary('__project_summary__','','Project Summary',leaves,0)]
    for index, entry in enumerate(items):
        item = entry.item
        level = len(item.wbs.split('.')) if item.wbs else 1
        if roots:
            level -= len(roots[0].item.wbs.split('.'))
        if item.role != 'item':
            members = [x for x in leaves if item.wbs and x.item.wbs.startswith(item.wbs + '.')]
            positions = [i for i,x in enumerate(items) if x in members]
            if any(i <= index for i in positions):
                raise ValueError('BOQ descendants precede their parent')
            if positions and any(not items[i].item.wbs.startswith(item.wbs + '.') for i in range(index+1,max(positions)+1)):
                raise ValueError('Non-contiguous BOQ branch')
            result.append(summary(item.key,item.wbs,item.description,members,max(level,0)))
        else:
            statuses = {p.status for p in entry.periods}
            status = 'Full' if statuses == {'Complete'} else '; '.join(sorted(statuses))
            if item.amount is None or item.amount < 0:
                status += '; ' + item.amount_status
            if any(p.warnings for p in entry.periods):status += '; Decreasing Requirement'
            if 'PartialMapping' in statuses:status += f' {entry.periods[0].mapping_coverage:.1%}'
            result.append(ReportRow(item.key,item.wbs,item.description,item.unit,item.quantity,item.amount,status,max(level,0),False,
                tuple(tuple(getattr(p,f) for p in entry.periods) for f in fields)))
    return tuple(result)


@dataclass(frozen=True)
class KnownProjectedTotals:
    cumulative: tuple[Decimal | None, ...]
    period: tuple[Decimal | None, ...]
    status: str


def known_projected_totals(projection):
    """Aggregate known leaf contributions only; never fill unknown money with zero.

    Opening known baseline is zero. Later period deltas require both endpoints
    of the same leaf contribution; a missing endpoint is excluded, not zeroed.
    Complete Project Summary remains independently governed by report_rows.
    """
    leaves = [entry for entry in projection.items if entry.item.role == 'item']
    keys = [entry.item.key for entry in leaves]
    if len(keys) != len(set(keys)):
        raise ValueError('Duplicate BOQ leaf identity')
    cumulative, period = [], []
    incomplete = False
    for index in range(len(projection.periods)):
        known, deltas = [], []
        for entry in leaves:
            row = entry.periods[index]
            amount = entry.item.amount
            if amount is None or amount < 0:
                incomplete = True
                continue
            current = row.known_amount
            previous = Decimal(0) if index == 0 else entry.periods[index - 1].known_amount
            if current is not None:
                known.append(current)
                if previous is not None:
                    deltas.append(current - previous)
            if amount > 0 and (row.cumulative_amount is None or current is None or previous is None):
                incomplete = True
        cumulative.append(sum(known, Decimal(0)) if known else None)
        period.append(sum(deltas, Decimal(0)) if deltas else None)
    status = ('Incomplete coverage — known contributions only' if incomplete
              else 'Known contributions — not certified payment')
    if not any(value is not None for value in cumulative):
        status = 'Unavailable — no determinable contributions'
    return KnownProjectedTotals(tuple(cumulative), tuple(period), status)


@dataclass(frozen=True)
class KnownProjectedPercentages:
    total_amount: Decimal | None
    cumulative: tuple[Decimal | None, ...]
    period: tuple[Decimal | None, ...]
    status: str


def known_projected_percentages(projection, known: KnownProjectedTotals):
    """Normalize existing known money by ALL distinct BOQ leaves, never subtotals."""
    leaves = [entry.item for entry in projection.items if entry.item.role == 'item']
    keys = [item.key for item in leaves]
    if len(keys) != len(set(keys)):
        raise ValueError('Duplicate BOQ leaf identity')
    valid = bool(leaves) and all(
        isinstance(item.amount, Decimal) and item.amount.is_finite() and item.amount >= 0
        for item in leaves
    )
    total = sum((item.amount for item in leaves), Decimal(0)) if valid else None
    reason = ('Invalid or missing total BOQ Amount' if total is None else
              'No positive total BOQ Amount' if total == 0 else '')
    def fractions(values):
        return tuple(value / total if not reason and value is not None else None
                     for value in values)
    return KnownProjectedPercentages(total, fractions(known.cumulative),
                                    fractions(known.period),
                                    known.status + ('; ' + reason if reason else ''))
