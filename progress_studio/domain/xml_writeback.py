"""XW-1 evidence models. No Excel coordinates, UI state or float conversion."""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, localcontext
import re


def monetary_decimal(raw: str) -> Decimal:
    """Keep stored digits, bounded to the existing Excel monetary range.

    A generous lexical limit prevents pathological exponents/coefficient sizes;
    accepted precision is never rounded to cents or converted through float.
    """
    text = raw.strip()
    if len(text) > 1024 or not re.fullmatch(
        r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]{1,4})?', text
    ):
        raise ValueError('missing or nonnumeric Amount')
    try:
        value = Decimal(text)
    except InvalidOperation:
        raise ValueError('invalid Amount') from None
    if not value.is_finite() or value < 0:
        raise ValueError('negative or nonfinite Amount')
    if value and not Decimal('2.2250738585072014e-308') <= value <= Decimal('9.99999999999999e307'):
        raise ValueError('Amount outside supported Excel numeric range')
    return value


def exact_total(values) -> Decimal:
    values = tuple(values)
    with localcontext() as ctx:
        ctx.prec = max(28, max((v.adjusted() for v in values), default=0)
                       - min((v.as_tuple().exponent for v in values), default=0)
                       + len(str(len(values))) + 4)
        return sum(values, Decimal(0))


@dataclass(frozen=True)
class WorkbookAmount:
    activity_id: str
    description: str
    row: int
    raw: str
    value: Decimal | None
    formula: bool
    milestone: bool | None


@dataclass(frozen=True)
class WorkbookAmounts:
    rows: tuple[WorkbookAmount, ...]
    errors: tuple[str, ...]


@dataclass(frozen=True)
class ActivityAmountChange:
    activity_id: str
    name: str
    source_row: int
    source_amount: Decimal
    before_raw: str | None
    after_raw: str | None
    status: str  # replace, unchanged, preserve_milestone
