"""Source-neutral, single-currency cash forecast records (no workbook metadata)."""
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from enum import Enum


class CashDirection(str, Enum):
    IN = 'in'
    OUT = 'out'


class CashCategory(str, Enum):
    OPERATING = 'operating'
    FINANCING = 'financing'


@dataclass(frozen=True)
class CashEvent:
    event_id: str
    cash_date: date
    amount: Decimal
    direction: CashDirection
    category: CashCategory = CashCategory.OPERATING
    item_id: str | None = None
    correction_reason: str = ''  # Required for negative amounts, which reverse this stream.


@dataclass(frozen=True)
class ForecastItem:
    item_id: str
    amount: Decimal
    direction: CashDirection
    expected_date: date | None = None
    category: CashCategory = CashCategory.OPERATING
    balance_date: date | None = None  # Amount is remaining at END of this date, if supplied.
    retention: bool = False


@dataclass(frozen=True)
class ForecastOverride:
    item_id: str
    reason: str
    remaining_amount: Decimal | None = None
    cash_date: date | None = None
    clear_date: bool = False


@dataclass(frozen=True)
class FinanceInputs:
    currency: str
    opening_date: date  # Opening cash is immediately BEFORE events on this day.
    opening_balance: Decimal
    actuals_through: date
    events: tuple[CashEvent, ...] = ()
    items: tuple[ForecastItem, ...] = ()
    overrides: tuple[ForecastOverride, ...] = ()
    actuals_complete: bool = False
    receivables_complete: bool = False
    cash_out_complete: bool = False
    unlinked_actuals_reconciled: bool = False
    horizon_end: date | None = None  # Extends reporting; never truncates cash events.


@dataclass(frozen=True)
class ResolvedForecastItem:
    item: ForecastItem
    settled_amount: Decimal
    base_remaining: Decimal
    effective_remaining: Decimal
    effective_date: date | None
    override_reason: str
    status: str  # settled / scheduled / overdue / undated


@dataclass(frozen=True)
class CashPoint:
    cash_date: date  # First day of month for monthly reporting points.
    actual_in: Decimal
    actual_out: Decimal
    forecast_in: Decimal
    forecast_out: Decimal
    closing_balance: Decimal


@dataclass(frozen=True)
class FinanceResult:
    currency: str
    items: tuple[ResolvedForecastItem, ...]
    dated_points: tuple[CashPoint, ...]
    monthly_points: tuple[CashPoint, ...]
    actual_in: Decimal  # Opening date through actuals_through, inclusive.
    actual_out: Decimal
    operating_cash_out_ac_proxy: Decimal  # All supplied history through actuals_through.
    remaining_receivable: Decimal  # Includes undated/overdue operating receipts.
    remaining_cash_out: Decimal  # Includes undated/overdue operating payments.
    remaining_retention: Decimal
    minimum_balance: Decimal
    minimum_balance_date: date
    peak_funding_requirement: Decimal
    issues: tuple[str, ...]

    @property
    def forecast_complete(self) -> bool:
        return not self.issues
