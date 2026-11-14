from dataclasses import replace
from decimal import Decimal as D
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
import pytest
from progress_studio.domain.boq_mapping_trace import mapping_trace
from progress_studio.domain.boq_payment_projection import project
from progress_studio.domain.boq_manual_requirements import compose, ManualInput
from progress_studio.domain.mapping_models import AllocationRecord
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from tests.unit.payment.test_boq_payment_projection import item, payment, run


def example():
    items=tuple(replace(item(key),wbs=f'1.{i+1}',description=f'BOQ {key}') for i,key in enumerate('abc'))
    p=project(items,(AllocationRecord('a','A',60),AllocationRecord('a','B',40),AllocationRecord('b','A',25)),payment(),('A','B','C'))
    p=replace(p,activity_names=(('A','Duplicate name'),('B','Duplicate name'),('C','No mapping')))
    return compose(p,(ManualInput('c',(D(0),None,D('.1')),9),))


def values(w,name):
    sheet,ref=next(w.defined_names[name].destinations)
    c,r,ec,er=range_boundaries(ref)
    return tuple(w[sheet].iter_rows(min_row=r,max_row=er,min_col=c,max_col=ec,values_only=True))


def test_snapshot_indexes_and_view_controls(tmp_path):
    p=example();w=render(p);path=tmp_path/'trace.xlsx';w.save(path);w.close();w=load_workbook(path)
    assert w.sheetnames==['Percentage','Amount','BOQ Mapping','Activity Mapping','_BP_Data']
    assert 'BP_TraceViews' not in w.defined_names
    for name, controls in (('BOQ Mapping', {'B2'}), ('Activity Mapping', {'B2'})):
        assert {str(d.sqref) for d in w[name].data_validations.dataValidation}==controls
        assert 'Trace view' not in [c.value for row in w[name] for c in row]
    assert 'Duplicate name' in w['Activity Mapping']['B2'].value
    edges=values(w,'BP_Relationships')
    assert [(r[0],r[2],r[4]) for r in edges]==[('a','A',.6),('a','B',.4),('b','A',.25)]
    for name in ('BP_RelBOQRows','BP_RelActivityRows'):
        assert sorted(row[0] for row in values(w,name))==[1,2,3]
    assert len(values(w,'BP_Activities'))==3
    assert 'Partial Mapping' in values(w,'BP_RelBOQ')[1][2]
    assert values(w,'BP_RelBOQ')[2][1]==0
    assert 'Manual BOQ' in values(w,'BP_RelBOQ')[2][2]
    assert w['_BP_Data'].sheet_state=='hidden' and w['_BP_Data'].protection.sheet
    assert not any(c.data_type=='f' for row in w['_BP_Data'] for c in row)
    for name in ('BOQ Mapping','Activity Mapping'):
        assert w[name].print_title_rows=='$1:$8' and w[name].freeze_panes=='B9'
        formulas=[c.value for row in w[name] for c in row if c.data_type=='f']
        assert all('BP_Periods' not in f and 'Payment Trace!' not in f and 'B3' not in f for f in formulas)
        assert w[name].print_area and w[name].page_setup.fitToWidth==1
    assert 'Payment Trace' not in w.sheetnames
    w.close()


@pytest.mark.parametrize('count',[1,84,90])
def test_reverse_capacity_is_not_hardcoded(count):
    items=tuple(replace(item(str(i)),wbs=f'1.{i+1}') for i in range(count))
    p=project(items,tuple(AllocationRecord(x.key,'A',100) for x in items),payment(),('A','B'))
    p=replace(p,activity_names=(('A','Activity'),('B','Unmapped')))
    assert mapping_trace(p).capacity==count
    w=render(p);assert w['Activity Mapping'].cell(8+count,3).data_type=='f'
    assert values(w,'BP_RelActivities')[0][4]==count
    assert w['Activity Mapping'].print_area.endswith(f'$D${8+count}')
    assert w['Activity Mapping']['A7'].value is None
    w.close()


def test_empty_mapping_and_legacy_missing_name_are_explicit():
    w=render(run(shares=()));assert values(w,'BP_RelBOQ')[0][1:]==(0,'Unmapped',0,0);w.close()
    w=render(run());assert values(w,'BP_Relationships')[0][3]=='Name unavailable';w.close()


def test_untrusted_names_remain_literals():
    p=replace(run(),activity_names=(('A','=HYPERLINK("bad")'),('B','Name')))
    w=render(p);assert values(w,'BP_Relationships')[0][3]=='=HYPERLINK("bad")'
    sheet,ref=next(w.defined_names['BP_Relationships'].destinations);_,r,_,_=range_boundaries(ref)
    assert w[sheet].cell(r,4).data_type=='s';w.close()


def test_cancel_in_mapping_snapshot():
    from openpyxl import Workbook
    from progress_studio.infrastructure.excel.boq_mapping_trace_workbook import prepare_mapping_trace
    w=Workbook();w.active.title='_BP_Data'
    with pytest.raises(InterruptedError):prepare_mapping_trace(w,run(),lambda:True)
    w.close()


@pytest.mark.parametrize('basis,expected', [
    (None, 'Unavailable'), (D('-1'), 'Unavailable'), (D(0), 0),
    (D('0.01'), 0.0025), (100, 25),
])
def test_allocated_money_uses_original_share_and_unrounded_basis(basis, expected):
    workbook = render(run(shares=(25,), amount=basis))
    edge, = values(workbook, 'BP_Relationships')
    assert edge[4:] == (.25, expected)
    assert values(workbook, 'BP_RelBOQ')[0][3:] == (.25, expected)
    assert values(workbook, 'BP_RelActivities')[0][5] == expected
    workbook.close()
