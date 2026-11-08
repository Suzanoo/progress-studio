from dataclasses import replace
from decimal import Decimal as D
import pytest
from openpyxl import load_workbook
from progress_studio.domain.boq_payment_report import known_projected_totals, known_projected_percentages, KnownProjectedTotals, report_rows
from progress_studio.domain.boq_payment_projection import project
from progress_studio.domain.mapping_models import AllocationRecord
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from tests.unit.payment.test_boq_payment_projection import run, item, payment


def normalized(result):
    return known_projected_percentages(result, known_projected_totals(result))


def test_fixed_all_leaf_denominator_and_signed_changes():
    a=item('a',D(100)); b=replace(item('b',D(200)),wbs='1.2')
    g=replace(item('g',D(99999),'group'),wbs='1')
    result=project((g,a,b),(AllocationRecord('a','A',100),),payment(a=(1,.5,0)),('A','B'))
    value=normalized(result)
    assert value.total_amount == 300
    assert value.cumulative == (D(100)/300,D(50)/300,D(0))
    assert value.period == (D(100)/300,-D(50)/300,-D(50)/300)
    assert 'Incomplete' in value.status
    assert report_rows(result)[0].values[0] == (None,)*3
    assert normalized(replace(result,items=result.items[1:]))==value


@pytest.mark.parametrize('amount,reason',[(None,'Invalid'),(D(-1),'Invalid'),(D('NaN'),'Invalid'),(D('Infinity'),'Invalid'),(D(0),'No positive')])
def test_invalid_or_zero_full_denominator_unavailable(amount,reason):
    result=run(); entry=result.items[0]
    result=replace(result,items=(replace(entry,item=replace(entry.item,amount=amount)),))
    value=known_projected_percentages(result,KnownProjectedTotals((D(10),)*3,(D(0),)*3,'Incomplete'))
    assert value.cumulative == (None,)*3 and value.period == (None,)*3
    assert reason in value.status and 'Incomplete' in value.status


def test_unknown_zero_partial_and_duplicate_identity():
    assert normalized(run(shares=())).cumulative == (None,)*3
    assert normalized(run(a=(0,0,0),b=(0,0,0))).cumulative == (D(0),)*3
    assert normalized(run(shares=(60,))).cumulative[-1] == D('.48')
    result=run()
    with pytest.raises(ValueError,match='Duplicate'):
        normalized(replace(result,items=result.items*2))
    unknown=KnownProjectedTotals((None,D(0),D(5)),(None,None,D(-5)),'Incomplete')
    value=known_projected_percentages(result,unknown)
    assert value.cumulative==(None,D(0),D('.05'))
    assert value.period==(None,None,D('-.05'))


def test_zero_leaf_does_not_replace_unknown_positive_leaf():
    a=item('a',D(0));b=replace(item('b',D(200)),wbs='1.2')
    result=project((a,b),(AllocationRecord('a','A',100),),payment(),('A','B'))
    value=normalized(result)
    assert value.total_amount==200 and value.cumulative==(D(0),)*3
    assert 'Incomplete' in value.status


def snapshot(w,name):
    sheet,address=next(w.defined_names[name].destinations)
    return [[c.value for c in row] for row in w[sheet][address]]


def test_row_five_selector_structure_and_immutable_money(tmp_path):
    w=render(run(a=(1,0,1),b=(1,1,1)));p=tmp_path/'percentage.xlsx';w.save(p);w.close();w=load_workbook(p)
    s=w['Percentage']
    assert s['A5'].value.startswith('Known Projected % of Total BOQ')
    assert 'Total BOQ Amount: 100.00' in s['A2'].value
    assert s.freeze_panes=='G8' and s.print_title_rows=='$1:$7' and s.print_title_cols=='$A:$F'
    assert 'A5:E5' in s.merged_cells and s.row_dimensions[5].outlineLevel==0
    assert s['G5'].number_format=='0.0%;[Red]-0.0%;0.0%'
    assert 'Invalid view' in s['G5'].value and 'Cumulative %' in s['G5'].value
    assert not s.protection.sheet and w['_BP_Data'].protection.sheet and w['_BP_Data'].sheet_state=='hidden'
    assert snapshot(w,'BP_KnownAmounts')==[[100,40,100],[100,-60,60]]
    assert snapshot(w,'BP_KnownPercentages')==[[1,.4,1],[1,-.6,.6]]
    assert snapshot(w,'BP_KnownCompletion')==[[1,0,1]]
    assert snapshot(w,'BP_TotalBOQAmount')==[[100]]
    assert all(c.data_type!='f' for row in w['_BP_Data'] for c in row)
    rules=[r for k,rs in s.conditional_formatting._cf_rules.items() if str(k.sqref) in ('H5','I5') for r in rs]
    assert len(rules)==2
    assert all('Cumulative %' in r.formula[0] and 'BP_KnownCompletion' in r.formula[0] and r.dxf.numFmt.formatCode==';;;' for r in rules)
    assert not any(str(k.sqref)=='G5' for k in s.conditional_formatting)
    w.close()


def test_near_completion_flags_are_decimal_exact():
    result=run(a=(1,1,1),b=(1,1,1));entry=result.items[0]
    almost=D('99.99999999999999999999999999')
    entry=replace(entry,periods=tuple(replace(p,known_amount=v) for p,v in zip(entry.periods,(D(100),almost,D(100)))))
    w=render(replace(result,items=(entry,)))
    assert snapshot(w,'BP_KnownCompletion')==[[1,0,1]]
    w.close()
