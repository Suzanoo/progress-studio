from dataclasses import replace
import os
import errno
from pathlib import Path

import pytest

from progress_studio.services.xml_writeback_service import (
    WritebackRequest, XmlWritebackService, WritebackCancelled,
)
from progress_studio.infrastructure.schedule_xml.p6_adapter import P6XmlAdapter
from progress_studio.services.import_service import ImportService
from tests.fixtures.xml_writeback import workbook, p6


@pytest.fixture
def operation(tmp_path):
    request=WritebackRequest(workbook(tmp_path/'source.xlsx'),p6(tmp_path/'target.xml'),
                             'p6:udf:47',monetary_confirmed=True)
    service=XmlWritebackService()
    preview=service.assess(request)
    assert preview.valid,preview.errors
    return service,preview,tmp_path/'output.xml'


def assert_originals(preview,originals,output):
    assert preview.request.workbook.read_bytes()==originals[0]
    assert preview.request.target_xml.read_bytes()==originals[1]
    assert not output.exists()
    assert not list(output.parent.glob('.progress-studio-xml-*.tmp'))


def test_review_required_and_tampered_preview_rejected(operation):
    service,a,output=operation
    with pytest.raises(ValueError,match='Review'):
        service.export(a,output)
    modified=replace(a,rows=tuple(replace(r,after_raw='9999') for r in a.rows))
    with pytest.raises(ValueError,match='modified'):
        service.export(modified,output,reviewed=True)
    assert not output.exists()


@pytest.mark.parametrize('which',['workbook','target_xml'])
def test_source_mutation_requires_new_review(operation,which):
    service,a,output=operation
    path=getattr(a.request,which)
    path.write_bytes(path.read_bytes()+b' ')
    with pytest.raises(ValueError,match='changed'):
        service.export(a,output,reviewed=True)
    assert not output.exists()


@pytest.mark.parametrize('checkpoint',[1,2,3])
def test_cancel_at_each_boundary_cleans_stage(operation,checkpoint):
    service,a,output=operation
    originals=(a.request.workbook.read_bytes(),a.request.target_xml.read_bytes())
    count=0
    def cancelled():
        nonlocal count
        count+=1
        return count==checkpoint
    with pytest.raises(WritebackCancelled):
        service.export(a,output,reviewed=True,cancelled=cancelled)
    assert_originals(a,originals,output)


def test_source_changes_after_stage_blocks_publication(operation):
    service,a,output=operation
    count=0
    def cancelled():
        nonlocal count
        count+=1
        if count==3:
            a.request.target_xml.write_bytes(a.request.target_xml.read_bytes()+b' ')
        return False
    with pytest.raises(ValueError,match='changed'):
        service.export(a,output,reviewed=True,cancelled=cancelled)
    assert not output.exists()
    assert not list(output.parent.glob('.progress-studio-xml-*.tmp'))


@pytest.mark.parametrize('operation_name',['mkstemp','fdopen','fsync','link'])
def test_io_failure_does_not_publish_partial_output(operation,monkeypatch,operation_name):
    service,a,output=operation
    originals=(a.request.workbook.read_bytes(),a.request.target_xml.read_bytes())
    import progress_studio.services.xml_writeback_service as module
    def fail(*args,**kwargs):
        raise OSError('injected write failure')
    owner=module.tempfile if operation_name=='mkstemp' else module.os
    monkeypatch.setattr(owner,operation_name,fail)
    with pytest.raises(OSError,match='injected'):
        service.export(a,output,reviewed=True)
    assert_originals(a,originals,output)


def test_corrupt_stage_is_detected(operation,monkeypatch):
    service,a,output=operation
    original=Path.read_bytes
    def read(path):
        data=original(path)
        return b'corrupt' if path.suffix=='.tmp' else data
    monkeypatch.setattr(Path,'read_bytes',read)
    with pytest.raises(OSError,match='integrity'):
        service.export(a,output,reviewed=True)
    assert not output.exists()
    assert not list(output.parent.glob('.progress-studio-xml-*.tmp'))


def test_destination_collision_before_and_during_publication(operation,monkeypatch):
    service,a,output=operation
    output.write_bytes(b'existing')
    with pytest.raises(FileExistsError):
        service.export(a,output,reviewed=True)
    assert output.read_bytes()==b'existing'
    output.unlink()
    real_link=os.link
    def race(src,dst):
        Path(dst).write_bytes(b'concurrent writer')
        real_link(src,dst)
    monkeypatch.setattr(os,'link',race)
    with pytest.raises(FileExistsError):
        service.export(a,output,reviewed=True)
    assert output.read_bytes()==b'concurrent writer'
    assert not list(output.parent.glob('.progress-studio-xml-*.tmp'))


def test_input_paths_and_aliases_are_never_outputs(operation):
    service,a,output=operation
    for path in (a.request.workbook,a.request.target_xml):
        with pytest.raises(ValueError,match='separate'):
            service.export(a,path,reviewed=True)
    os.link(a.request.target_xml,output)
    with pytest.raises(FileExistsError):
        service.export(a,output,reviewed=True)


def test_missing_parent_is_safe(operation):
    service,a,output=operation
    with pytest.raises(FileNotFoundError):
        service.export(a,output.parent/'missing'/'output.xml',reviewed=True)


def _symlink_or_skip(link, target):
    # Catch only setup capability errors, never exceptions from the engine.
    try:
        link.symlink_to(target)
    except OSError as exc:
        if getattr(exc, 'winerror', None) == 1314 or exc.errno in {
            errno.EPERM, errno.EACCES, errno.ENOSYS, errno.ENOTSUP,
        }:
            pytest.skip(f'Symlink creation unavailable in test directory: {exc}')
        raise


def test_input_symlink_alias_is_never_output(operation):
    service,a,output=operation
    _symlink_or_skip(output,a.request.target_xml)
    with pytest.raises(ValueError,match='separate'):
        service.export(a,output,reviewed=True)
    assert output.is_symlink()


def test_dangling_symlink_is_safe(operation):
    service,a,output=operation
    _symlink_or_skip(output,output.parent/'missing-file')
    with pytest.raises(FileExistsError):
        service.export(a,output,reviewed=True)
    assert output.is_symlink()


def test_cleanup_failure_reports_completed_output_truthfully(operation,monkeypatch):
    service,a,output=operation
    real=Path.unlink
    def fail(path,*args,**kwargs):
        if path.suffix=='.tmp':
            raise OSError('cleanup injected')
        return real(path,*args,**kwargs)
    monkeypatch.setattr(Path,'unlink',fail)
    result=service.export(a,output,reviewed=True)
    assert output.is_file()
    assert any('cleanup required' in w for w in result.warnings)
    monkeypatch.undo()
    for path in output.parent.glob('.progress-studio-xml-*.tmp'):
        path.unlink()


def test_existing_p6_import_and_create_amount_semantics(operation):
    service,a,output=operation
    service.export(a,output,reviewed=True)
    schedule=P6XmlAdapter().normalize(output)
    rows={r.activity_id:r for r in schedule.activities}
    assert dict(rows['A1'].amount_field_values)['p6:udf:47']=='12.34567890123456789'
    assert dict(rows['M1'].amount_field_values)['p6:udf:47']=='500.50'
    # Supported Create import path intentionally gives milestones zero weight.
    from progress_studio.infrastructure.schedule_xml import NormalizedScheduleXmlReader
    from progress_studio.services.schedule_service import ScheduleService
    from progress_studio.infrastructure.excel import ImportWorkbookWriter
    from openpyxl import load_workbook
    importer=ImportService(NormalizedScheduleXmlReader(),ScheduleService(),ImportWorkbookWriter())
    imported=output.with_suffix('.xlsx')
    assert importer.import_xml(output,imported,weight_basis='amount',amount_field='p6:udf:47')[2]==4
    wb=load_workbook(imported,data_only=True)
    try:
        # Verify monetary milestone retained independently in the accepted EV input.
        records=list(wb['EV Monetary Inputs'].iter_rows(min_row=4,values_only=True))
        values={r[0]:r[2] for r in records}
        assert values['M1']==500.5
        assert values['A1']==pytest.approx(float('12.34567890123456789'))
    finally:
        wb.close()


def test_partial_write_failure_cleans_incomplete_stage(operation,monkeypatch):
    service,a,output=operation
    real=os.fdopen
    class BrokenStream:
        def __init__(self,stream): self.stream=stream
        def __enter__(self): return self
        def __exit__(self,*args): self.stream.close()
        def write(self,payload):
            self.stream.write(payload[:100])
            raise OSError('disk full after partial write')
    monkeypatch.setattr(os,'fdopen',lambda *args,**kwargs:BrokenStream(real(*args,**kwargs)))
    with pytest.raises(OSError,match='disk full'):
        service.export(a,output,reviewed=True)
    assert not output.exists()
    assert not list(output.parent.glob('.progress-studio-xml-*.tmp'))


def test_proof_harness_requires_exact_review_digest(operation,capsys):
    import runpy
    service,a,output=operation
    module=runpy.run_path(str(Path(__file__).resolve().parents[3]/'scripts/proofs/xml_amount_writeback.py'))
    args=['--workbook',str(a.request.workbook),'--target-xml',str(a.request.target_xml),
          '--field','p6:udf:47','--monetary-confirmed']
    assert module['main'](args)==0
    text=capsys.readouterr().out
    digest=text.split('Review digest: ')[1].strip()
    assert not output.exists()
    assert module['main'](args+['--review-digest','wrong','--output',str(output)])==2
    assert not output.exists()
    assert module['main'](args+['--review-digest',digest,'--output',str(output)])==0
    assert output.exists()


@pytest.mark.parametrize('code', [1314, errno.EPERM, errno.ENOTSUP])
def test_symlink_capability_denial_is_narrowly_skipped(tmp_path, monkeypatch, code):
    error = OSError('simulated unavailable symlink capability')
    if code == 1314:
        error.winerror = 1314
    else:
        error.errno = code
    def denied(*args, **kwargs):
        raise error
    monkeypatch.setattr(Path, 'symlink_to', denied)
    with pytest.raises(pytest.skip.Exception, match='Symlink creation unavailable'):
        _symlink_or_skip(tmp_path/'link', tmp_path/'target')


def test_unexpected_symlink_setup_error_is_not_hidden(tmp_path, monkeypatch):
    def failed(*args, **kwargs):
        raise OSError(errno.EIO, 'unexpected I/O failure')
    monkeypatch.setattr(Path, 'symlink_to', failed)
    with pytest.raises(OSError, match='unexpected I/O failure'):
        _symlink_or_skip(tmp_path/'link', tmp_path/'target')
