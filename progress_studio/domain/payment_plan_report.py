"""Read-only Payment Plan Report values. All fractions use 1 for 100%."""
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, localcontext
from types import MappingProxyType
from typing import Mapping

from progress_studio.domain.payment_requirement_states import RequirementState

ZERO = Decimal(0)
COMPLETE = Decimal('0.9999')


@dataclass(frozen=True)
class Period:
    id: str
    eligible: date | None
    reference: date | None
    issues: tuple[str, ...]


@dataclass(frozen=True)
class Row:
    key: str
    parent: str | None
    code: str
    name: str
    source_row: int
    activities: tuple[str, ...]
    activity: bool


@dataclass(frozen=True)
class Leaf:
    amount: Decimal | None
    schedule: tuple[RequirementState, ...]


@dataclass(frozen=True)
class Value:
    fraction: Decimal | None
    delta: Decimal | None
    money: Decimal | None
    movement: Decimal | None
    known_money: Decimal | None
    known_basis: Decimal
    known_leaves: int
    display_fraction: Decimal | None
    transition: str
    correction: bool
    period_visible: bool
    cumulative_visible: bool


@dataclass(frozen=True)
class Scope:
    activities: tuple[str, ...]
    basis: Decimal | None
    values: tuple[Value, ...]


@dataclass(frozen=True)
class Selection:
    revision: str
    rows: tuple[Row, ...]
    included_in: Mapping[str, str]
    row_values: Mapping[str, Scope]
    selected: Scope
    unselected: Scope
    project: Scope
    full_project: bool
    retained_money: tuple[Decimal, ...]
    visible_known_money: tuple[Decimal | None, ...]


@dataclass(frozen=True)
class ReportSource:
    project: str
    filename: str
    revision: str
    monetary_origin: str
    periods: tuple[Period, ...]
    rows: tuple[Row, ...]
    leaves: Mapping[str, Leaf]
    diagnostics: tuple[str, ...]

    def __post_init__(self):
        object.__setattr__(self, 'leaves', MappingProxyType(dict(self.leaves)))

    def select(self, keys, *, revision: str):
        if revision != self.revision:
            raise ValueError('Workbook revision changed. Review the selection again.')
        wanted = set(keys)
        by_key = {row.key: row for row in self.rows}
        missing = wanted - by_key.keys()
        if missing:
            raise ValueError(f'Unknown reporting rows: {sorted(missing)}')
        rows = tuple(row for row in self.rows if row.key in wanted)
        overlap = {}
        covered = set()
        for row in rows:
            covered.update(row.activities)
            ancestor = row.parent
            while ancestor is not None and ancestor not in wanted:
                ancestor = by_key[ancestor].parent
            if ancestor is not None:
                overlap[row.key] = ancestor
        with localcontext() as ctx:
            ctx.prec = 80
            values = {r.key: self._scope(r.activities, r.activity) for r in rows}
            selected = self._scope(tuple(a for a in self.leaves if a in covered))
            other = self._scope(tuple(a for a in self.leaves if a not in covered))
            project = self._scope(tuple(self.leaves))
            retained = tuple(sum((values[r.key].values[j].money for r in rows
                                  if r.key not in overlap and
                                  values[r.key].values[j].transition == 'suppressed'), ZERO)
                             for j in range(len(self.periods)))
            visible_known = tuple(
                None if v.known_money is None else v.known_money - kept
                for v, kept in zip(selected.values, retained))
        return Selection(revision, rows, MappingProxyType(overlap), MappingProxyType(values),
                         selected, other, project,
                         len(covered) == len(self.leaves), retained, visible_known)

    def _scope(self, ids, activity=False):
        leaves = [self.leaves[a] for a in ids]
        basis = None if any(v.amount is None for v in leaves) else sum(
            (v.amount for v in leaves), ZERO)
        points = []
        for j in range(len(self.periods)):
            known = [(v.amount, v.schedule[j].fraction) for v in leaves
                     if v.amount is not None and v.schedule[j].fraction is not None]
            known_basis = sum((amount for amount, _ in known), ZERO)
            subtotal = sum((amount * fraction for amount, fraction in known), ZERO)
            money = subtotal if basis is not None and known_basis == basis else None
            fraction = leaves[0].schedule[j].fraction if activity else (
                money / basis if basis and money is not None else None)
            points.append((fraction, money, subtotal if known or not ids else None, known_basis, len(known)))
        return Scope(tuple(ids), basis, tuple(series(points)))


def difference(current, previous):
    return None if current is None or previous is None else current - previous


def series(points):
    previous = money_before = ZERO
    completed = False
    for fraction, money, known, basis, count in points:
        delta = difference(fraction, previous)
        movement = difference(money, money_before)
        valid = fraction is not None and money is not None
        done = valid and fraction >= COMPLETE
        was_done = previous is not None and previous >= COMPLETE
        changed = delta not in (None, ZERO) or movement not in (None, ZERO)
        state = 'open'
        if not valid:
            state = 'unresolved'
        elif done and was_done:
            state = 'adjusted_complete' if changed else 'suppressed'
        elif done:
            state = 'recompleted' if completed else 'completed'
        elif was_done:
            state = 'reopened'
        yield Value(fraction, delta, money, movement, known, basis, count,
                    Decimal(1) if done else fraction, state,
                    (delta or ZERO)<0 or (movement or ZERO)<0,
                    state!='suppressed' and (not valid or changed),
                    state!='suppressed' and (not valid or fraction!=ZERO or changed))
        previous, money_before = fraction, money
        completed |= done
