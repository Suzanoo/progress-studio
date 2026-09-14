"""Pure cash calculations. Contract Value is not a cost budget or accounting EAC."""
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext

from progress_studio.domain.financial_forecast import (
    CashCategory, CashDirection, CashPoint, FinanceInputs, FinanceResult,
    ForecastItem, ResolvedForecastItem,
)

ZERO = Decimal('0.00')
CENT = Decimal('0.01')


class FinanceValidationError(ValueError):
    pass


def money(value, *, signed=False):
    """V1 uses two decimal places, half-up; float conversion uses its decimal text."""
    try:
        if isinstance(value, bool):
            raise ValueError
        amount = Decimal(str(value))
        if not amount.is_finite() or abs(amount) > Decimal('1e18'):
            raise ValueError
        if not signed and amount < 0:
            raise ValueError
        return amount.quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise FinanceValidationError('Invalid monetary amount') from exc


def _date(value):
    if type(value) is not date:
        raise FinanceValidationError('Expected a calendar date without time')
    return value


def _identity(value):
    if not isinstance(value, str) or not value.strip() or value != value.strip():
        raise FinanceValidationError('IDs must be nonblank and trimmed')


def _unique(records, key):
    result = {}
    for record in records:
        value = getattr(record, key)
        _identity(value)
        if value in result:
            raise FinanceValidationError(f'Duplicate {key}: {value}')
        result[value] = record
    return result


class FinancialForecastDeriver:
    def derive(self, inputs: FinanceInputs) -> FinanceResult:
        # Isolate precision/rounding from the caller's Decimal context.
        with localcontext() as context:
            context.prec = 50
            return self._derive(inputs)

    def _derive(self, inputs):
        _identity(inputs.currency)
        _date(inputs.opening_date)
        _date(inputs.actuals_through)
        if inputs.opening_date > inputs.actuals_through:
            raise FinanceValidationError('Opening date is after actuals cutoff')
        if inputs.horizon_end is not None:
            _date(inputs.horizon_end)
        opening = money(inputs.opening_balance, signed=True)
        items = _unique(inputs.items, 'item_id')
        events = _unique(inputs.events, 'event_id')
        overrides = _unique(inputs.overrides, 'item_id')
        if overrides.keys() - items.keys():
            raise FinanceValidationError('Override references unknown item')
        issues = []
        for flag in ('actuals_complete', 'receivables_complete', 'cash_out_complete',
                     'unlinked_actuals_reconciled'):
            if type(getattr(inputs, flag)) is not bool:
                raise FinanceValidationError(f'{flag} must be boolean')
        for flag in ('actuals_complete', 'receivables_complete', 'cash_out_complete'):
            if not getattr(inputs, flag):
                issues.append(flag + ':not_confirmed')
        for item in items.values():
            money(item.amount)
            if not isinstance(item.direction, CashDirection) or not isinstance(item.category, CashCategory):
                raise FinanceValidationError('Invalid item direction/category')
            if type(item.retention) is not bool or (item.retention and (
                    item.direction != CashDirection.IN or item.category != CashCategory.OPERATING)):
                raise FinanceValidationError('Retention must be an operating receivable')
            if item.expected_date is not None:
                _date(item.expected_date)
            if item.balance_date is not None and _date(item.balance_date) > inputs.actuals_through:
                raise FinanceValidationError('Remaining balance snapshot is after cutoff')
        buckets = defaultdict(lambda: [ZERO, ZERO, ZERO, ZERO])
        settlements = defaultdict(lambda: ZERO)
        proxy = ZERO
        for event_id, event in sorted(events.items()):
            _date(event.cash_date)
            if not isinstance(event.direction, CashDirection) or not isinstance(event.category, CashCategory):
                raise FinanceValidationError('Invalid cash direction/category')
            amount = money(event.amount, signed=True)
            if Decimal(str(event.amount)) < 0 and not event.correction_reason.strip():
                raise FinanceValidationError('Negative cash requires a correction reason')
            if event.item_id is not None:
                if event.item_id not in items:
                    raise FinanceValidationError('Settlement references unknown item')
                item = items[event.item_id]
                if (item.direction, item.category) != (event.direction, event.category):
                    raise FinanceValidationError('Settlement direction/category mismatch')
            if event.cash_date > inputs.actuals_through:
                issues.append(f'future_actual:{event_id}')
                continue
            if event.item_id is None and amount and not inputs.unlinked_actuals_reconciled:
                issues.append(f'unlinked_actual:{event_id}')
            if event.item_id is not None:
                item = items[event.item_id]
                if item.balance_date is None or event.cash_date > item.balance_date:
                    settlements[event.item_id] += amount
            if event.direction == CashDirection.OUT and event.category == CashCategory.OPERATING:
                proxy += amount
            if event.cash_date >= inputs.opening_date:
                buckets[event.cash_date][0 if event.direction == CashDirection.IN else 1] += amount
        resolved = []
        remaining_in = remaining_out = retention = ZERO
        for item_id, item in sorted(items.items()):
            settled = settlements[item_id]
            base = money(item.amount) - settled
            if settled < 0 or base < 0:
                raise FinanceValidationError(f'Settlement does not reconcile: {item_id}')
            effective, cash_date, reason = base, item.expected_date, ''
            override = overrides.get(item_id)
            if override:
                if not isinstance(override.reason, str) or not override.reason.strip():
                    raise FinanceValidationError('Override requires a reason')
                if type(override.clear_date) is not bool or (override.clear_date and override.cash_date is not None):
                    raise FinanceValidationError('Conflicting override dates')
                if override.remaining_amount is not None:
                    effective = money(override.remaining_amount)
                if override.cash_date is not None:
                    cash_date = _date(override.cash_date)
                if override.clear_date:
                    cash_date = None
                reason = override.reason
            status = ('settled' if effective == 0 else 'undated' if cash_date is None
                      else 'overdue' if cash_date <= inputs.actuals_through else 'scheduled')
            resolved.append(ResolvedForecastItem(item, settled, base, effective, cash_date, reason, status))
            if item.category == CashCategory.OPERATING:
                if item.direction == CashDirection.IN:
                    remaining_in += effective
                else:
                    remaining_out += effective
            if item.retention:
                retention += effective
            if status in ('undated', 'overdue'):
                issues.append(f'{status}:{item_id}')
            elif status == 'scheduled':
                buckets[cash_date][2 if item.direction == CashDirection.IN else 3] += effective
        balance = minimum = opening
        minimum_date = inputs.opening_date
        points = []
        monthly = defaultdict(lambda: [ZERO, ZERO, ZERO, ZERO])
        for day, values in sorted(buckets.items()):
            balance += values[0] - values[1] + values[2] - values[3]
            points.append(CashPoint(day, *values, balance))
            if balance < minimum:
                minimum, minimum_date = balance, day
            totals = monthly[day.replace(day=1)]
            for index, value in enumerate(values):
                totals[index] += value
        end = max(inputs.actuals_through, inputs.horizon_end or inputs.actuals_through,
                  max(buckets, default=inputs.actuals_through))
        month = inputs.opening_date.replace(day=1)
        month_end = end.replace(day=1)
        balance = opening
        monthly_points = []
        while month <= month_end:
            values = monthly[month]
            balance += values[0] - values[1] + values[2] - values[3]
            monthly_points.append(CashPoint(month, *values, balance))
            if month == month_end:
                break
            month = date(month.year + (month.month == 12), month.month % 12 + 1, 1)
        return FinanceResult(inputs.currency, tuple(resolved), tuple(points), tuple(monthly_points),
                             sum((p.actual_in for p in points), ZERO),
                             sum((p.actual_out for p in points), ZERO), proxy,
                             remaining_in, remaining_out, retention, minimum, minimum_date,
                             max(ZERO, -minimum), tuple(issues))


def certificate_items(item_id, gross, certification_date, *, credit_days=30,
                      retention_rate=Decimal('0.05'), retention_release_date=None):
    """Gross eligible contract certificate -> net receivable + explicit retention."""
    _identity(item_id)
    _date(certification_date)
    if type(credit_days) is not int or credit_days < 0:
        raise FinanceValidationError('Credit days must be a nonnegative integer')
    if retention_release_date is not None:
        _date(retention_release_date)
    with localcontext() as context:
        context.prec = 50
        amount = money(gross)
        try:
            rate = Decimal(str(retention_rate))
            if isinstance(retention_rate, bool) or not rate.is_finite() or not 0 <= rate <= 1:
                raise ValueError
        except (InvalidOperation, ValueError) as exc:
            raise FinanceValidationError('Retention rate must be between zero and one') from exc
        retained = money(amount * rate)
        try:
            receipt_date = certification_date + timedelta(days=credit_days)
        except (OverflowError, ValueError) as exc:
            raise FinanceValidationError('Receipt date out of range') from exc
        return (ForecastItem(item_id + ':net', amount - retained, CashDirection.IN, receipt_date),
                ForecastItem(item_id + ':retention', retained, CashDirection.IN,
                             retention_release_date, retention=True))


def remaining_cash_out(total_outturn, paid_to_date, known_outstanding):
    """Uncommitted residual only; explicit cash estimate, never BAC/CPI."""
    with localcontext() as context:
        context.prec = 50
        residual = money(total_outturn) - money(paid_to_date) - money(known_outstanding)
        if residual < 0:
            raise FinanceValidationError('Cash outturn is below paid plus known outstanding')
        return residual


def phase_remaining(amount, weighted_dates, *, progress_through):
    """Allocate a separately reconciled remaining amount to supplied future dates.

    Largest-remainder cents; date breaks ties. No schedule inference or credit shift.
    """
    _date(progress_through)
    with localcontext() as context:
        context.prec = 50
        amount = money(amount)
        weights = {}
        for day, value in weighted_dates:
            _date(day)
            if day <= progress_through or day in weights:
                raise FinanceValidationError('Phasing requires unique future dates')
            try:
                weight = Decimal(str(value))
                if isinstance(value, bool) or not weight.is_finite() or weight < 0 or weight > Decimal('1e18'):
                    raise ValueError
            except (InvalidOperation, ValueError) as exc:
                raise FinanceValidationError('Invalid phasing weight') from exc
            weights[day] = weight
        total = sum(weights.values(), Decimal(0))
        if amount == 0:
            return ()
        if total == 0:
            raise FinanceValidationError('Positive remainder requires revised timing/weights')
        cents = int(amount / CENT)
        shares = [(day, Decimal(cents) * weight / total) for day, weight in sorted(weights.items()) if weight]
        allocations = {day: int(share) for day, share in shares}
        spare = cents - sum(allocations.values())
        for day, _ in sorted(shares, key=lambda pair: (-(pair[1] - int(pair[1])), pair[0]))[:spare]:
            allocations[day] += 1
        return tuple((day, Decimal(value) * CENT) for day, value in sorted(allocations.items()))


def remaining_unperformed_value(contract_value, certified_gross, earned_uncertified):
    """Disjoint Contract Value buckets; certified includes paid and retained gross.

    Caller confirms the real commercial basis. This does not infer certification
    from EV or generate a second copy of already-certified receivables.
    """
    with localcontext() as context:
        context.prec = 50
        remaining = money(contract_value) - money(certified_gross) - money(earned_uncertified)
        if remaining < 0:
            raise FinanceValidationError('Contract Value buckets do not reconcile')
        return remaining
