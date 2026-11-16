from decimal import Decimal
from openpyxl import load_workbook
import pytest

from progress_studio.infrastructure.excel.payment_plan_adjustment import create_adjustment, read_adjustment
from tests.unit.payment.test_payment_plan_report import source


def test_reviewed_plan_roundtrip_and_completion(tmp_path):
    s = source(['.2','.3','.4','.5','.6','.7'], ['1']*6)
    selection = s.select(['W','A'],revision=s.revision)
    path = tmp_path/'plan.xlsx'
    create_adjustment(path,s,selection)
    assert read_adjustment(path,s).selection == selection
    with pytest.raises(FileExistsError):
        create_adjustment(path,s,selection)
    w = load_workbook(path)
    for i,x in enumerate([.99989999,.9999,.9999,.99991,.4,1],5):
        w['Edit plan'].cell(4,i,x)
    w.save(path)
    plan = read_adjustment(path,s)
    assert [v.transition for v in plan.selection.selected.values] == [
        'open','completed','suppressed','adjusted_complete','reopened','recompleted']
    assert plan.issues and not plan.incomplete
    assert plan.selection.retained_money[2]==Decimal('999.9')
    w['Edit plan']['E4']=None
    w['Edit plan']['F4']=0
    w.save(path)
    with pytest.raises(ValueError,match='changed'):
        plan.validate()
    incomplete=read_adjustment(path,s)
    assert incomplete.selection.selected.values[1].money==0
    assert incomplete.selection.selected.values[1].movement is None
    with pytest.raises(ValueError,match='Resolve'):
        incomplete.validate()


@pytest.mark.parametrize('value',['=50%',2,-.1,'50%'])
def test_invalid_cell_rejected(tmp_path,value):
    s=source(['.5'],['.7'])
    path=tmp_path/'plan.xlsx'
    create_adjustment(path,s,s.select(['A'],revision=s.revision))
    w=load_workbook(path)
    w['Edit plan']['E4']=value
    w.save(path)
    with pytest.raises(ValueError,match='E4'):
        read_adjustment(path,s)
