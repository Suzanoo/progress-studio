from datetime import datetime
from dataclasses import replace

from openpyxl import load_workbook

from tests.fixtures.boq_revision import revision_case
from tests.integration.mapping.test_mapping_overlay_export import _created_mapping_input
from progress_studio.domain.mapping_models import AllocationRecord
from progress_studio.infrastructure.excel.mapping_reader import ProgressActivityReader, BOQSheetReader
from progress_studio.infrastructure.excel.amount_workbook import find_header
from progress_studio.infrastructure.session.mapping_session_repository import materialize_workbook_snapshot
from progress_studio.services.mapping_store import MappingStore
from progress_studio.services.workbook_export_service import WorkbookExportService


def test_updated_project_exports_new_amounts_preserving_workbook_inputs(tmp_path,monkeypatch):
    c=revision_case(tmp_path)
    source,_=_created_mapping_input(tmp_path,False)
    w=load_workbook(source);header,h=find_header(w['main'],['Row Type','Activity ID','Amount'])
    row=next(r for r in range(header+1,w['main'].max_row+1) if w['main'].cell(r,h['activity id']).value=='A1000' and w['main'].cell(r,h['p/a']).value=='P')
    w['main'].cell(row+1,h['actual start'],datetime(2026,3,6))
    w['main'].cell(row+1,h['actual finish'],datetime(2026,3,13))
    protected={name:tuple(w[name].values) for name in ('EV Monetary Inputs','Payment Input') if name in w.sheetnames}
    assert 'EV Monetary Inputs' in protected
    w.save(source);w.close();source_bytes=source.read_bytes()
    store=MappingStore();store.load_activities(ProgressActivityReader().read(source))
    store.load_boq(BOQSheetReader().read(c.old,'Project'))
    store.restore_allocations([AllocationRecord('Project|2|2','A1000',100),AllocationRecord('Project|3|3','A1000',25)])
    session=c.repo.create(source,c.old,'Project',store.allocation_records(),working_tree_nodes=list(store.working_tree_nodes()))
    c.repo.save(c.project,session)
    monkeypatch.setattr(c.repo,'materialize_snapshot',lambda s:materialize_workbook_snapshot(s,tmp_path/'cache'))
    p=c.service.preview(session,store,c.project,c.new)
    staged,_,_=c.service.apply(p,session,store)
    reopened=c.repo.load(c.project)
    assert reopened.allocations==session.allocations
    output=tmp_path/'updated-progress.xlsx'
    WorkbookExportService().export(source,output,staged)
    assert source.read_bytes()==source_bytes
    w=load_workbook(output);header,h=find_header(w['main'],['Row Type','Activity ID','Amount'])
    row=next(r for r in range(header+1,w['main'].max_row+1) if w['main'].cell(r,h['activity id']).value=='A1000' and w['main'].cell(r,h['p/a']).value=='P')
    assert w['main'].cell(row,h['amount']).value==2000
    assert w['main'].cell(row+1,h['actual start']).value==datetime(2026,3,6)
    assert w['main'].cell(row+1,h['actual finish']).value==datetime(2026,3,13)
    assert w['main'].cell(row+1,h['actual duration']).data_type=='f'
    for name,values in protected.items():
        actual=list(w[name].values)
        if name=='Payment Input':
            # Reconciliation refreshes D5 and PI-1 Amount snapshots. All
            # requirements and helper formulas must otherwise be preserved.
            status=list(actual[4]);status[3]=values[4][3];actual[4]=tuple(status)
            for i, record in enumerate(actual):
                if record[0] in {'ACT', 'WBS'}:
                    assert record[4] == 2000
                    normalized=list(record);normalized[4]=values[i][4]
                    actual[i]=tuple(normalized)
        assert tuple(actual)==values
    assert w['main'].cell(row,h['plan duration']).data_type=='f'
    assert 'BOQ Activity Mapping' in w.sheetnames
    w.close()
