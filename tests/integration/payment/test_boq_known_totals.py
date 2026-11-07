from dataclasses import replace
from decimal import Decimal as D
import pytest
from openpyxl import load_workbook
from progress_studio.domain.boq_payment_report import known_projected_totals, report_rows
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from tests.unit.payment.test_boq_payment_projection import run, item, payment
from progress_studio.domain.boq_payment_projection import project
from progress_studio.domain.mapping_models import AllocationRecord


def test_known_complete_and_signed_changes():
    totals = known_projected_totals(run(a=(1, 0, None), b=(1, 1, 1)))
    assert totals.cumulative == (D(100), D(40), D(40))
    assert totals.period == (D(100), D(-60), D(0))
    assert 'not certified' in totals.status


@pytest.mark.parametrize('shares,b', [((60,), (None,.5,1)), ((60,40), (None,None,None))])
def test_partial_known_is_not_complete(shares, b):
    result = run(shares=shares, b=b)
    assert known_projected_totals(result).cumulative[-1] == D(48)
    assert 'Incomplete' in known_projected_totals(result).status
    assert report_rows(result)[0].values[2][-1] is None


@pytest.mark.parametrize('amount,shares', [(None,(60,40)),(D(-1),(60,40)),(D(100),())])
def test_unavailable_is_not_numeric_zero(amount, shares):
    totals = known_projected_totals(run(amount=amount,shares=shares))
    assert totals.cumulative == (None,)*3 and totals.period == (None,)*3
    assert 'Unavailable' in totals.status


def test_zero_preserved():
    totals = known_projected_totals(run(amount=D(0)))
    assert totals.cumulative == (D(0),)*3 and totals.period == (D(0),)*3


def test_distinct_leaf_only_and_unmapped_warning():
    a=item('a',D(100)); b=replace(item('b',D(200)),wbs='1.2'); g=replace(item('g',D(9999),'group'),wbs='1')
    result=project((g,a,b),(AllocationRecord('a','A',100),),payment(),('A','B'))
    totals=known_projected_totals(result)
    assert totals.cumulative == (D(20),D(20),D(80))
    assert totals.period == (D(20),D(0),D(60))
    assert 'Incomplete' in totals.status
    with pytest.raises(ValueError,match='Duplicate'):
        known_projected_totals(replace(result,items=result.items+(result.items[1],)))


def test_missing_prior_endpoint_not_zeroed():
    result=run(); entry=result.items[0]
    entry=replace(entry,periods=(replace(entry.periods[0],known_amount=None),)+entry.periods[1:])
    totals=known_projected_totals(replace(result,items=(entry,)))
    assert totals.period[:2] == (None,None) and 'Incomplete' in totals.status


def test_frozen_printed_formula_row_and_snapshot(tmp_path):
    w=render(run(a=(1,0,None),b=(1,1,1)));p=tmp_path/'known.xlsx';w.save(p);w.close()
    w=load_workbook(p);s=w['Amount']
    assert s['A5'].value.startswith('Known Projected Amount')
    assert s.freeze_panes=='G8' and s.print_title_rows=='$1:$7'
    assert s.row_dimensions[5].outlineLevel==0 and not s.protection.sheet
    assert 'Invalid view' in s['G5'].value and 'Cumulative Amount' in s['G5'].value
    assert s['G5'].number_format.endswith(';0.00')
    data=w['_BP_Data']; assert [data.cell(data.max_row-1,c).value for c in range(1,4)]==[100,40,40]
    assert [data.cell(data.max_row,c).value for c in range(1,4)]==[100,-60,0]
    assert 'A5:E5' in s.merged_cells and len(s.data_validations.dataValidation)==1
    w.close()
