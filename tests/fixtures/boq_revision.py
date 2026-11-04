from pathlib import Path
from types import SimpleNamespace
from openpyxl import Workbook, load_workbook
from progress_studio.domain.mapping_models import ActivityRow, AllocationRecord
from progress_studio.infrastructure.excel.mapping_reader import BOQSheetReader
from progress_studio.infrastructure.session import MappingSessionRepository
from progress_studio.services.mapping_store import MappingStore
from progress_studio.services.boq_update_service import BOQUpdateService


def revision_case(tmp_path: Path):
    old, new = tmp_path/'old.xlsx', tmp_path/'new.xlsx'
    w = Workbook(); s = w.active; s.title = 'Project'
    s.append(['WBS No.', 'WBS-2', 'WBS-3', 'WBS-4', 'Description', 'Unit', 'Qty', 'Amount'])
    s.append(['1.1', 'Building', 'Interior', '', 'Ceiling', 'm2', 20, 1000])
    s.append(['1.2', 'Building', 'Equipment', '', 'Lift', 'item', 1, 2000])
    s.append(['1.3', 'Building', 'Equipment', '', 'Electrical', 'item', 1, 500])
    s.append(['1.4', 'Building', 'Excluded', '', 'Zero item', 'item', 1, 0])
    s.append(['2', 'Building', 'Heading', '', 'Heading', None, None, None])
    w.save(old);s['H2'] = 1500;s['H4'] = 600;w.save(new);w.close()
    progress = tmp_path/'progress.xlsx';w=Workbook();w.save(progress);w.close()
    store=MappingStore();store.load_activities([ActivityRow('A1','1','1.1','First'),ActivityRow('A2','1','1.2','Second')])
    store.load_boq(BOQSheetReader().read(old,'Project'))
    store.restore_allocations([AllocationRecord('Project|2|2','A1',60),AllocationRecord('Project|2|2','A2',40),AllocationRecord('Project|3|3','A2',95)])
    repo=MappingSessionRepository()
    session=repo.create(progress,old,'Project',store.allocation_records(),working_tree_nodes=list(store.working_tree_nodes()))
    project=tmp_path/'case.progressstudio';repo.save(project,session)
    return SimpleNamespace(old=old,new=new,progress=progress,store=store,repo=repo,session=session,project=project,service=BOQUpdateService(repo))


def edit(path, mutate):
    w=load_workbook(path);mutate(w['Project']);w.save(path);w.close()


def cache_formula(path, coordinate, formula, cached, cell_type=None):
    """Build XLSX formula-result fixtures without relying on an Excel engine."""
    import zipfile
    import xml.etree.ElementTree as ET
    ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with zipfile.ZipFile(path) as z:parts={i.filename:z.read(i.filename) for i in z.infolist()}
    root=ET.fromstring(parts['xl/worksheets/sheet1.xml'])
    cell=next((c for c in root.iter(ns+'c') if c.attrib['r']==coordinate),None)
    if cell is None:
        number=''.join(c for c in coordinate if c.isdigit())
        row=next(r for r in root.iter(ns+'row') if r.get('r')==number)
        cell=ET.SubElement(row,ns+'c',r=coordinate)
    for child in list(cell):cell.remove(child)
    cell.attrib.pop('t',None)
    if cell_type:cell.set('t',cell_type)
    ET.SubElement(cell,ns+'f').text=formula
    if cached is not None:ET.SubElement(cell,ns+'v').text=str(cached)
    parts['xl/worksheets/sheet1.xml']=ET.tostring(root)
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)
