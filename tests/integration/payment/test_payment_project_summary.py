"""Overall Payment summary: source metadata, legacy fallback and numeric proof."""
import pytest
from openpyxl import load_workbook
from progress_studio.infrastructure.excel.payment_input_summary import project_summary_tree
from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook
from progress_studio.services.payment_service import PaymentService
from tests.fixtures.payment_requirements import requirement_source, edit_requirements
from tests.integration.payment.test_payment_requirement_summary import values, activity_rows, recalculate

@pytest.mark.parametrize('embedded',[True,False])
@pytest.mark.parametrize('metadata',[True,False])
def test_project_metadata_and_preservation(tmp_path,embedded,metadata):
    source=requirement_source(tmp_path/'source.xlsx')
    w=load_workbook(source);s=w['main'];s['O14']=50
    if metadata:
        for c,v in {1:'Project Summary',2:'SOURCE-ROOT',3:'Actual Project',4:'P',8:0}.items():s.cell(5,c,v)
    else:
        s['A5']=None; s['D5']=None
    w.save(source);w.close()
    output=tmp_path/'out.xlsx'
    if embedded:PaymentService().prepare_embedded_payment_input(source,output,3)
    else:PaymentInputWorkbook().create(source,output,3)
    edit_requirements(output);before=values(output)
    w=load_workbook(output);s=w['Payment Input']
    assert s['A8'].value=='WBS' and s['C8'].value is None
    assert s['D8'].value==('Actual Project' if metadata else 'Project Summary')
    assert s['B8'].value==('SOURCE-ROOT' if metadata else None)
    assert s['E8'].value==650
    assert s.row_dimensions[8].outlineLevel==0
    assert all(s.row_dimensions[r].outlineLevel>0 for r in range(9,s.max_row+1))
    assert all(s.cell(8,c).protection.locked for c in range(1,9))
    assert s.freeze_panes=='F8' and s['B9'].value=='1'
    w.close()
    if embedded:
        PaymentService().prepare_embedded_payment_input(output,output,3)
        assert values(output)==before
        w=load_workbook(output);s=w['Payment Input']
        assert s.max_row==20 and s['E8'].value==650;w.close()


def test_project_numerical_states(tmp_path):
    source=requirement_source(tmp_path/'source.xlsx')
    w=load_workbook(source);w['main']['O14']=50;w.save(source);w.close()
    output=tmp_path/'cases.xlsx'
    PaymentService().prepare_embedded_payment_input(source,output,3);edit_requirements(output)
    w=load_workbook(output);base=w['Payment Input']
    for s in list(w):
        if s!=base:w.remove(s)
    cases={'complete':{},'unknown':{'A200':(None,None,None)},'invalid':{'A200':(.2,'bad',None)},
           'zero':{a:(0,None,None) for a in activity_rows(base)},
           'allblank':{a:(None,None,None) for a in activity_rows(base)},
           'decrease_return':{a:(1,.5,1) for a in activity_rows(base)}}
    for name,updates in cases.items():
        s=w.copy_worksheet(base);s.title=name;rows=activity_rows(s)
        for a,seq in updates.items():
            for c,v in enumerate(seq,6):s.cell(rows[a],c).value=v
    bad=w.copy_worksheet(base);bad.title='bad_amount';bad['E8']='Check Amount'
    w.save(output);w.close()
    calc=recalculate(output,tmp_path/'calc');w=load_workbook(calc,data_only=True)
    assert [w['complete'].cell(8,c).value for c in range(6,9)]==pytest.approx([285/650,185/650,185/650])
    for name,text in [('unknown','Incomplete'),('invalid','Check input'),('bad_amount','Check Amount')]:
        assert all(w[name].cell(8,c).value==text for c in range(6,9))
    assert [w['zero'].cell(8,c).value for c in range(6,9)]==[0,0,0]
    assert all(w['allblank'].cell(8,c).value in ('',None) for c in range(6,9))
    assert [w['decrease_return'].cell(8,c).value for c in range(6,9)]==[1,.5,1]
    w.close()


def test_existing_genuine_root_and_legacy_level_zero():
    def row(kind,level,name):return dict(row_type=kind,outline_level=level,wbs=name,activity_name=name,activity_id='A')
    genuine=[row('WBS',0,'ROOT'),row('WBS',1,'1'),row('ACT',2,'1.1')]
    result=project_summary_tree(genuine)
    assert len(result)==3 and result[0]['wbs']=='ROOT' and result[0]['activity_id']==''
    separate=[row('WBS',0,'1'),row('ACT',1,'1.1'),row('WBS',0,'2'),row('ACT',1,'2.1')]
    result=project_summary_tree(separate)
    assert len(result)==5 and result[0]['wbs']==''
    assert [r['outline_level'] for r in result]==[0,1,2,1,2]
    assert separate[0]['outline_level']==0
