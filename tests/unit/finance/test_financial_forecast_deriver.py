from dataclasses import FrozenInstanceError, replace
from datetime import date, datetime
from decimal import Decimal, localcontext

import pytest

from progress_studio.domain.financial_forecast import (
    CashCategory as Category, CashDirection as Direction, CashEvent, FinanceInputs,
    ForecastItem, ForecastOverride,
)
from progress_studio.services.financial_forecast_deriver import (
    FinanceValidationError, FinancialForecastDeriver, certificate_items,
    money, phase_remaining, remaining_cash_out,
)

D = Decimal
JAN = date(2026, 1, 1)
CUT = date(2026, 1, 31)
MAR = date(2026, 3, 2)


def inputs(**kwargs):
    values = dict(currency='THB', opening_date=JAN, opening_balance=D(0),
                  actuals_through=CUT, actuals_complete=True,
                  receivables_complete=True, cash_out_complete=True,
                  unlinked_actuals_reconciled=True)
    values.update(kwargs)
    return FinanceInputs(**values)


def derive(**kwargs):
    return FinancialForecastDeriver().derive(inputs(**kwargs))


def test_certificate_calendar_days_and_partial_settlement():
    items = certificate_items('certificate', 100, CUT)
    assert [i.amount for i in items] == [D(95), D(5)]
    assert items[0].expected_date == MAR
    result = derive(items=items, events=(CashEvent('receipt', CUT, D(40), Direction.IN,
                                                item_id='certificate:net'),))
    assert result.remaining_receivable == D(60)
    assert result.remaining_retention == D(5)
    assert result.items[0].base_remaining == D(55)
    assert result.actual_in == D(40)
    assert result.monthly_points[-1].forecast_in == D(55)
    assert result.issues == ('undated:certificate:retention',)


def test_retention_release_is_explicit_and_not_withheld_twice():
    release = date(2027, 1, 1)
    result = derive(items=certificate_items('c', 100, CUT, retention_release_date=release),
                    horizon_end=date(2026, 6, 1))
    assert result.monthly_points[-1].cash_date == release
    assert result.monthly_points[-1].closing_balance == D(100)
    assert result.forecast_complete


def test_net_manual_receipt_gets_no_default_terms():
    result = derive(items=(ForecastItem('net', D(95), Direction.IN, MAR),))
    assert result.remaining_receivable == D(95)
    assert result.remaining_retention == 0


def test_snapshot_does_not_subtract_prior_settlements_again():
    item = ForecastItem('invoice', D(60), Direction.IN, MAR, balance_date=date(2026, 1, 15))
    result = derive(items=(item,), events=(
        CashEvent('before', date(2026, 1, 10), D(40), Direction.IN, item_id='invoice'),
        CashEvent('after', date(2026, 1, 20), D(10), Direction.IN, item_id='invoice')))
    assert result.items[0].settled_amount == D(10)
    assert result.remaining_receivable == D(50)
    assert result.actual_in == D(50)


def test_opening_cash_excludes_prior_history_but_proxy_includes_it():
    result = derive(opening_date=date(2026, 1, 15), opening_balance=D(100), events=(
        CashEvent('old', JAN, D(20), Direction.OUT),
        CashEvent('today', date(2026, 1, 15), D(10), Direction.OUT)))
    assert result.actual_out == D(10)
    assert result.operating_cash_out_ac_proxy == D(30)
    assert result.monthly_points[-1].closing_balance == D(90)


def test_financing_separate_from_ac_proxy_and_operating_remaining():
    result = derive(events=(CashEvent('loan', CUT, D(100), Direction.IN, Category.FINANCING),
                            CashEvent('fee', CUT, D(10), Direction.OUT, Category.FINANCING),
                            CashEvent('cost', CUT, D(20), Direction.OUT)),
                    items=(ForecastItem('repay', D(90), Direction.OUT, MAR, Category.FINANCING),))
    assert result.operating_cash_out_ac_proxy == D(20)
    assert result.remaining_cash_out == 0
    assert result.peak_funding_requirement == D(20)


@pytest.mark.parametrize('amount,expected', [(None, 60), (0, 0), (75, 75)])
def test_override_replaces_remaining_not_actual(amount, expected):
    result = derive(items=(ForecastItem('i', D(100), Direction.IN, MAR),),
                    events=(CashEvent('paid', CUT, D(40), Direction.IN, item_id='i'),),
                    overrides=(ForecastOverride('i', 'Revised collection', amount),))
    assert result.items[0].base_remaining == 60
    assert result.items[0].effective_remaining == expected
    assert result.actual_in == 40
    assert result.items[0].item.amount == 100


def test_overdue_undated_and_future_actual_are_visible_not_silently_shifted():
    result = derive(items=(ForecastItem('late', D(100), Direction.IN, CUT),
                          ForecastItem('unknown', D(50), Direction.OUT)),
                    events=(CashEvent('future', MAR, D(100), Direction.IN, item_id='late'),))
    assert result.remaining_receivable == 100
    assert result.remaining_cash_out == 50
    assert result.actual_in == 0
    assert result.issues == ('future_actual:future', 'overdue:late', 'undated:unknown')
    assert not result.forecast_complete


def test_override_can_reschedule_or_explicitly_clear_date():
    item = ForecastItem('i', D(100), Direction.IN, CUT)
    assert derive(items=(item,), overrides=(ForecastOverride('i', 'Expected date', cash_date=MAR),)).forecast_complete
    result = derive(items=(item,), overrides=(ForecastOverride('i', 'Unknown', clear_date=True),))
    assert result.items[0].effective_date is None
    assert result.items[0].item.expected_date == CUT


def test_correction_reopens_remaining_and_reverses_cash():
    result = derive(items=(ForecastItem('i', D(100), Direction.OUT, MAR),), events=(
        CashEvent('pay', JAN, D(70), Direction.OUT, item_id='i'),
        CashEvent('refund', CUT, D(-20), Direction.OUT, item_id='i', correction_reason='Returned payment')))
    assert result.actual_out == 50
    assert result.operating_cash_out_ac_proxy == 50
    assert result.remaining_cash_out == 50


def test_completeness_requires_explicit_declarations_and_unlinked_reconciliation():
    result = derive(actuals_complete=False, receivables_complete=False, cash_out_complete=False,
                    unlinked_actuals_reconciled=False,
                    events=(CashEvent('e', CUT, D(1), Direction.IN),))
    assert len(result.issues) == 4
    assert 'unlinked_actual:e' in result.issues


def test_empty_confirmed_forecast_includes_opening_deficit_and_empty_months():
    result = derive(opening_balance=D(-50), horizon_end=MAR)
    assert result.peak_funding_requirement == 50
    assert result.minimum_balance_date == JAN
    assert len(result.monthly_points) == 3
    assert all(p.closing_balance == -50 for p in result.monthly_points)


def test_cash_out_residual_reconciles_known_and_uncommitted():
    assert remaining_cash_out(1000, 250, 400) == 350
    with pytest.raises(FinanceValidationError):
        remaining_cash_out(600, 250, 400)


def test_phasing_conserves_cents_and_uses_future_progress_cutoff():
    dates = [(date(2026, 2, day), 1) for day in (1, 2, 3)]
    allocation = phase_remaining('100.00', reversed(dates), progress_through=CUT)
    assert [v for _, v in allocation] == [D('33.34'), D('33.33'), D('33.33')]
    assert sum(v for _, v in allocation) == 100
    assert phase_remaining(0, [], progress_through=CUT) == ()


@pytest.mark.parametrize('weights', [[], [(MAR, 0)], [(JAN, 1)], [(MAR, -1)],
                                     [(MAR, 'NaN')], [(MAR, True)], [(MAR, 1), (MAR, 2)]])
def test_invalid_or_missing_phasing_is_not_hidden(weights):
    with pytest.raises(FinanceValidationError):
        phase_remaining(100, weights, progress_through=CUT)


@pytest.mark.parametrize('value', ['NaN', 'Infinity', '-Infinity', True, -1, 'invalid', '1e30'])
def test_invalid_money(value):
    with pytest.raises(FinanceValidationError):
        money(value)


def test_rounding_and_caller_precision_do_not_change_service_results():
    with localcontext() as context:
        context.prec = 3
        items = certificate_items('i', '100000.01', CUT)
        result = derive(items=items)
    assert result.remaining_receivable == D('100000.01')
    assert money('1.005') == D('1.01')
    assert money(0.1) == D('0.10')


@pytest.mark.parametrize('change', [
    dict(opening_date=MAR), dict(actuals_through=datetime(2026, 1, 31)),
    dict(currency=' '), dict(actuals_complete='yes'),
    dict(items=(ForecastItem('i', D(-1), Direction.IN),)),
    dict(items=(ForecastItem('i', D(1), Direction.IN, balance_date=MAR),)),
    dict(items=(ForecastItem('i', D(1), Direction.OUT, retention=True),)),
    dict(events=(CashEvent('e', JAN, D(-1), Direction.IN),)),
    dict(events=(CashEvent('e', JAN, D(1), Direction.IN, item_id='missing'),)),
    dict(events=(CashEvent('e', JAN, D(1), 'in'),)),
    dict(overrides=(ForecastOverride('missing', 'reason'),)),
])
def test_invalid_inputs_fail_before_producing_a_projection(change):
    with pytest.raises(FinanceValidationError):
        derive(**change)


@pytest.mark.parametrize('kind', ['item', 'event', 'override'])
def test_duplicate_stable_ids_rejected(kind):
    item = ForecastItem('i', D(1), Direction.IN, MAR)
    event = CashEvent('e', JAN, D(1), Direction.IN)
    override = ForecastOverride('i', 'reason')
    changes = {'item': dict(items=(item, item)), 'event': dict(events=(event, event)),
               'override': dict(items=(item,), overrides=(override, override))}
    with pytest.raises(FinanceValidationError):
        derive(**changes[kind])


@pytest.mark.parametrize('override', [ForecastOverride('i', ''), ForecastOverride('i', 'x', -1),
                                     ForecastOverride('i', 'x', cash_date=MAR, clear_date=True)])
def test_invalid_override(override):
    with pytest.raises(FinanceValidationError):
        derive(items=(ForecastItem('i', D(1), Direction.IN, MAR),), overrides=(override,))


@pytest.mark.parametrize('amount,direction', [(101, Direction.IN), (1, Direction.OUT), (-1, Direction.IN)])
def test_settlement_must_reconcile(amount, direction):
    with pytest.raises(FinanceValidationError):
        derive(items=(ForecastItem('i', D(100), Direction.IN, MAR),),
               events=(CashEvent('e', JAN, D(amount), direction, item_id='i', correction_reason='Correction'),))


@pytest.mark.parametrize('kwargs', [dict(credit_days=-1), dict(credit_days=1.5), dict(retention_rate=1.1),
                                    dict(retention_rate='NaN'), dict(retention_rate=True)])
def test_invalid_customer_terms(kwargs):
    with pytest.raises(FinanceValidationError):
        certificate_items('i', 100, CUT, **kwargs)


def test_input_records_are_not_mutated():
    data = inputs(items=(ForecastItem('i', D(100), Direction.IN, MAR),))
    original = replace(data)
    assert FinancialForecastDeriver().derive(data) == FinancialForecastDeriver().derive(data)
    assert data == original
    with pytest.raises(FrozenInstanceError):
        data.currency = 'USD'


def test_contract_value_buckets_exclude_certified_and_earned_uncertified():
    from progress_studio.services.financial_forecast_deriver import remaining_unperformed_value
    assert remaining_unperformed_value(1000, 400, 100) == 500
    with pytest.raises(FinanceValidationError):
        remaining_unperformed_value(400, 400, 1)


def test_opening_retention_and_partial_release_reconcile():
    item = ForecastItem('retention', D(50), Direction.IN, MAR, balance_date=JAN, retention=True)
    result = derive(items=(item,), events=(CashEvent('release', CUT, D(20), Direction.IN,
                                                    item_id='retention'),))
    assert result.remaining_retention == 30
    assert result.remaining_receivable == 30
    assert result.actual_in == 20
