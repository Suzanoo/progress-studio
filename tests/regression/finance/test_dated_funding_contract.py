from datetime import date
from decimal import Decimal as D
from dataclasses import replace

from progress_studio.domain.financial_forecast import CashDirection as Direction, CashEvent, FinanceInputs, ForecastItem
from progress_studio.services.financial_forecast_deriver import FinancialForecastDeriver


def test_month_end_surplus_must_not_hide_midmonth_funding_need():
    data = FinanceInputs('THB', date(2026, 1, 1), D(10), date(2026, 1, 31),
                         items=(ForecastItem('pay', D(100), Direction.OUT, date(2026, 2, 5)),
                                ForecastItem('receive', D(120), Direction.IN, date(2026, 2, 25))))
    result = FinancialForecastDeriver().derive(data)
    assert result.peak_funding_requirement == 90
    assert result.minimum_balance_date == date(2026, 2, 5)
    assert result.monthly_points[-1].closing_balance == 30


def test_same_day_netting_is_order_independent_not_an_intraday_claim():
    day = date(2026, 1, 15)
    events = (CashEvent('pay', day, D(100), Direction.OUT), CashEvent('receive', day, D(100), Direction.IN))
    data = FinanceInputs('THB', date(2026, 1, 1), D(0), date(2026, 1, 31), events=events)
    service = FinancialForecastDeriver()
    result = service.derive(data)
    assert result == service.derive(replace(data, events=tuple(reversed(events))))
    assert result.peak_funding_requirement == 0
    assert len(result.dated_points) == 1


def test_advancing_cutoff_moves_settlement_from_forecast_to_actual_once():
    day = date(2026, 2, 5)
    data = FinanceInputs('THB', date(2026, 1, 1), D(0), date(2026, 1, 31),
                         events=(CashEvent('received', day, D(40), Direction.IN, item_id='invoice'),),
                         items=(ForecastItem('invoice', D(100), Direction.IN, date(2026, 3, 1)),))
    service = FinancialForecastDeriver()
    before = service.derive(data)
    after = service.derive(replace(data, actuals_through=day))
    assert (before.actual_in, before.remaining_receivable) == (0, 100)
    assert (after.actual_in, after.remaining_receivable) == (40, 60)
    assert before.monthly_points[-1].closing_balance == after.monthly_points[-1].closing_balance == 100
