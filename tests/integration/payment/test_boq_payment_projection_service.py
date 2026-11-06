from pathlib import Path
from dataclasses import asdict
from decimal import Decimal
import json
import pytest
from openpyxl import Workbook,load_workbook
from tests.fixtures.payment_requirements import requirement_source,edit_requirements
from progress_studio.services.payment_service import PaymentService
from progress_studio.services.boq_payment_projection import BOQPaymentProjectionService
from progress_studio.infrastructure.excel.boq_projection_catalog import read_catalog
from progress_studio.infrastructure.session.mapping_session_repository import MappingSessionRepository
from progress_studio.domain.mapping_models import AllocationRecord


def fixture(tmp_path):
    progress=requirement_source(tmp_path/'source.xlsx');boq=tmp_path/'boq.xlsx'
    w=load_workbook(progress);mapping=w.create_sheet('Amount Mapping')
    mapping.append(['Activity ID','WBS','Description','Amount','Status'])
    for row in w['main'].iter_rows(min_row=6,values_only=True):
        if row[0]=='Activity' and row[3]=='P':mapping.append([row[4],row[1],row[2],row[14],'Mapped'])
    w.save(progress);w.close()
    w=Workbook();s=w.active;s.title='Project'
    s.append(['WBS No.','WBS-2','WBS-3','WBS-4','Description','Unit','Qty','Amount'])
    s.append(['1','Work',None,None,'Work',None,None,None])
    s.append(['1.1','Work','Floor',None,'Concrete','m2',10,100])
    s.append(['1.2','Work','Floor',None,'Other','m2',10,50])
    w.save(boq);w.close()
    repository=MappingSessionRepository()
    session=repository.create(progress_file=progress,boq_file=boq,boq_sheet='Project',allocations=[AllocationRecord('Project|3|3','A100',100)])
    path=tmp_path/'project.progressstudio';repository.save(path,session)
    current=tmp_path/'current.xlsx';PaymentService().prepare_embedded_payment_input(progress,current,3);edit_requirements(current)
    return path,current,boq


def test_service_keeps_sources_and_diagnostics_independent(tmp_path):
    session,current,boq=fixture(tmp_path);before={p:p.read_bytes() for p in (session,current,boq)}
    service=BOQPaymentProjectionService()
    with pytest.raises(ValueError,match='accept revision'):service.calculate(session,current)
    result=service.calculate(session,current,accept_revision=True,boq_path=boq)
    assert len(result.items)==3 and result.items[0].item.role=='group'
    assert result.items[1].periods[0].cumulative_amount==100
    assert result.items[2].periods[0].status=='Unmapped'
    assert 'SourceRevision' in {d.code for d in result.diagnostics}
    assert all(p.read_bytes()==data for p,data in before.items())
    w=load_workbook(current);w['main']['O7']=150;w.save(current);w.close()
    result=service.calculate(session,current,accept_revision=True)
    differences=[d for d in result.diagnostics if d.code=='ActivityAmountDifference']
    assert len(differences)==1 and 'difference=50' in differences[0].detail
    assert result.items[1].periods[0].cumulative_amount==100


@pytest.mark.parametrize('failure',['share','snapshot','unknown','description','boq_revision','duplicate_main','input'])
def test_failure_simulation_does_not_write_sources(tmp_path,failure):
    session,current,boq=fixture(tmp_path)
    if failure in ('share','snapshot','unknown'):
        payload=json.loads(session.read_text())
        if failure=='share':payload['allocations'][0]['share_percent']=float('nan')
        if failure=='snapshot':payload['boq_snapshot']['content_b64']='corrupt'
        if failure=='unknown':payload['allocations'][0]['activity_id']='MISSING'
        session.write_text(json.dumps(payload))
    elif failure=='boq_revision':
        w=load_workbook(boq);w['Project']['H3']=200;w.save(boq);w.close()
    else:
        w=load_workbook(current)
        if failure=='description':w['main']['C7']='Different item'
        if failure=='duplicate_main':w['main']['E9']='A100'
        if failure=='input':w['Payment Input']['F10']='bad'
        w.save(current);w.close()
    before={p:p.read_bytes() for p in (session,current,boq)}
    with pytest.raises(ValueError):BOQPaymentProjectionService().calculate(session,current,accept_revision=True,boq_path=boq)
    assert all(p.read_bytes()==data for p,data in before.items())


def test_lossless_catalog_and_unresolved_formula(tmp_path):
    _,_,boq=fixture(tmp_path)
    w=load_workbook(boq);s=w['Project']
    for n,value in enumerate((0,-1,None,'bad','=1+1',"=IF('Project'!H3=\"\",\"\",'Project'!H4)"),3):
        s.append([f'1.{n}','Work','Repeated',None,f'Item {n}','m',1,value])
    w.save(boq);w.close();rows=read_catalog(boq,'Project')
    assert len(rows)==9
    assert [r.amount_status for r in rows[-6:]]==['ZeroAmount','NegativeAmount','MissingAmount','InvalidAmount','UnresolvedFormula','UnresolvedFormula']


def test_saved_mapping_label_is_not_current_main_identity(tmp_path):
    session,current,_=fixture(tmp_path)
    payload=json.loads(session.read_text(encoding='utf-8'))
    # Mapping labels may have been edited independently of the source main.
    from progress_studio.infrastructure.excel.mapping_reader import ProgressActivityReader
    from progress_studio.services.mapping_store import MappingStore
    store=MappingStore();store.load_activities(ProgressActivityReader().read(tmp_path/'source.xlsx'))
    from dataclasses import replace
    nodes=tuple(replace(node,name='Mapping-specific label') if node.source_activity_id=='A100' else node for node in store.working_tree_nodes())
    repo=MappingSessionRepository();saved=repo.load(session)
    repo.save(session,replace(saved,working_tree_nodes=nodes))
    result=BOQPaymentProjectionService().calculate(session,current,accept_revision=True)
    assert result.items[1].periods[0].cumulative_amount==100
    assert any(d.code=='MappingDescriptionRevision' and d.identity=='A100' for d in result.diagnostics)
