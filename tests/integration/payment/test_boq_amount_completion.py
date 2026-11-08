from dataclasses import replace
from decimal import Decimal as D
import pytest
from openpyxl import load_workbook
from progress_studio.domain.boq_payment_report import amount_completion_states, report_rows
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from progress_studio.domain.boq_payment_projection import project
from progress_studio.domain.mapping_models import AllocationRecord
from tests.unit.payment.test_boq_payment_projection import run, item, payment


def flags(p):
    return amount_completion_states(p,report_rows(p))


@pytest.mark.parametrize('values,expected',[
 ((.9,1,1),(False,True,True)),((1,.9,1),(True,False,True)),
 ((.99999999,1,1),(False,True,True)),((0,0,0),(False,False,False)),
 ((None,None,None),(False,False,False)),((1,None,None),(True,True,True))])
def test_exact_state_for_leaf_and_project(values,expected):
    p=run(a=values,b=values)
    assert flags(p)==(expected,expected)


def test_rounded_float_cannot_create_completion():
    p=run(a=(1,1,1),b=(1,1,1));e=p.items[0]
    near=D('.9999999999999999999999999999')
    periods=(replace(e.periods[0],cumulative_fraction=near),)+e.periods[1:]
    p=replace(p,items=(replace(e,periods=periods),))
    assert float(near)==1 and flags(p)[-1]==(False,True,True)


@pytest.mark.parametrize('amount,shares',[(D(0),(60,40)),(None,(60,40)),(D(-1),(60,40)),(D(100),()),(D(100),(60,))])
def test_no_completion_for_zero_unknown_invalid_partial(amount,shares):
    assert not any(any(row) for row in flags(run(a=(1,1,1),b=(1,1,1),amount=amount,shares=shares)))


def test_nested_group_and_unmapped_positive_leaf():
    g=replace(item('g',D(9999),'group'),wbs='1')
    nested=replace(item('h',D(9999),'group'),wbs='1.1')
    a=replace(item('a'),wbs='1.1.1');b=replace(item('b'),wbs='1.2')
    p=project((g,nested,a,b),(AllocationRecord('a','A',100),),payment(a=(1,1,1)),('A','B'))
    states=flags(p)
    assert states==((False,)*3,(True,)*3,(True,)*3,(False,)*3)
    zero=replace(b,amount=D(0))
    p=project((g,nested,a,zero),(AllocationRecord('a','A',100),),payment(a=(1,1,1)),('A','B'))
    assert flags(p)[0]==(True,)*3


def test_saved_rules_do_not_change_values_or_known_headline(tmp_path):
    p=run(a=(1,.9,1),b=(1,.9,1));w=render(p);out=tmp_path/'amount.xlsx';w.save(out);w.close();w=load_workbook(out)
    sheet,address=next(w.defined_names['BP_AmountCompletion'].destinations)
    assert [[c.value for c in row] for row in w[sheet][address]]==[[1,0,1],[1,0,1]]
    s=w['Amount'];assert s['G5'].value.startswith('=IF') and 'Cumulative Amount' in s['G5'].value
    assert s['G8'].value.startswith('=IF')
    assert {str(k.sqref) for k in s.conditional_formatting}=={'H8:H9','I8:I9'}
    for rules in s.conditional_formatting._cf_rules.values():
        for rule in rules:
            assert 'Cumulative Amount' in rule.formula[0] and 'ROW()-7' in rule.formula[0]
            assert rule.dxf.numFmt.formatCode==';;;'
    assert s.freeze_panes=='G8' and s.print_title_rows=='$1:$7'
    assert not s.protection.sheet and w[sheet].protection.sheet
    w.close()
