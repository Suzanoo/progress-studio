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
