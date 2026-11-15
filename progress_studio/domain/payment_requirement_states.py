"""Shared cumulative target schedule for one Activity, in period order."""
from decimal import Decimal
from typing import NamedTuple


class RequirementState(NamedTuple):
    fraction: Decimal | None
    origin: str
    row: int | None
    column: int | None
    raw: Decimal | None
    decrease_seen: bool


def requirement_states(periods, entries):
    """Entries map period IDs to (explicit fraction, source row, source column)."""
    latest = Decimal(0)
    row = column = None
    started = decreased = False
    for period in periods:
        raw = None
        if period in entries:
            raw, row, column = entries[period]
            decreased |= raw < latest
            latest = raw
            started = True
            origin = 'Explicit'
        elif started:
            origin = 'Carried'
        else:
            origin = 'BeforeFirst' if entries else 'Unknown'
        yield RequirementState(latest if entries else None, origin, row, column, raw, decreased)
