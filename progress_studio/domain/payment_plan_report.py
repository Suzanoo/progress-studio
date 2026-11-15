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
        previous_fraction = previous_money = ZERO
        completed_before = False
        values = []
        for j in range(len(self.periods)):
            known = [(v.amount, v.schedule[j].fraction) for v in leaves
                     if v.amount is not None and v.schedule[j].fraction is not None]
            known_basis = sum((amount for amount, _ in known), ZERO)
            subtotal = sum((amount * fraction for amount, fraction in known), ZERO)
            money = subtotal if basis is not None and known_basis == basis else None
            fraction = leaves[0].schedule[j].fraction if activity else (
                money / basis if basis and money is not None else None)
            delta = difference(fraction, previous_fraction)
            movement = difference(money, previous_money)
            valid = fraction is not None and money is not None
            complete = valid and fraction >= COMPLETE
            previously_complete = previous_fraction is not None and previous_fraction >= COMPLETE
            correction = (delta is not None and delta < 0) or (movement is not None and movement < 0)
            changed = delta not in (None, ZERO) or movement not in (None, ZERO)
            if not valid:
                transition = 'unresolved'
            elif complete and previously_complete:
                transition = 'adjusted_complete' if changed else 'suppressed'
            elif complete:
                transition = 'recompleted' if completed_before else 'completed'
            elif previously_complete:
                transition = 'reopened'
            else:
                transition = 'open'
            visible = transition != 'suppressed'
            values.append(Value(fraction, delta, money, movement, subtotal if known or not ids else None,
                                known_basis, len(known), Decimal(1) if complete else fraction,
                                transition, correction, visible and (not valid or changed),
                                visible and (not valid or fraction != ZERO or changed)))
            completed_before |= complete
            previous_fraction, previous_money = fraction, money
        return Scope(tuple(ids), basis, tuple(values))


def difference(current, previous):
    return None if current is None or previous is None else current - previous
