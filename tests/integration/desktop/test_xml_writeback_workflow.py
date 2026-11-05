"""Real service + headless Desktop callbacks. Not native Windows/P6 acceptance."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import inspect
import threading
import tkinter as tk

import pytest

from progress_studio.presentation.gui import xml_writeback_dialog as ui
from progress_studio.presentation.gui.app import ProgressStudioDesktopApp
from progress_studio.presentation.gui import strings
from progress_studio.services.xml_writeback_service import XmlWritebackService
from tests.fixtures.xml_writeback import workbook, p6


class Control:
    def __init__(self, variable=None):
        self.options = {}; self.variable = variable; self.visible = False
    def configure(self, **kw): self.options.update(kw)
    def pack(self, **kw): self.visible = True
    def pack_forget(self): self.visible = False
    def start(self): self.running = True
    def stop(self): self.running = False
    def current(self):
        try: return self.options.get('values', ()).index(self.variable.get())
        except ValueError: return -1


class Tree:
    def __init__(self): self.rows = {}
    def get_children(self): return tuple(self.rows)
    def delete(self, *ids):
        for i in ids: self.rows.pop(i)
    def insert(self, parent, where, values): self.rows[str(len(self.rows))] = values
    def selection(self): return ('0',) if self.rows else ()
    def item(self, item, key): return self.rows[item]


@pytest.fixture
def dialog(monkeypatch):
    # Plain variables avoid creating Tcl interpreters in headless worker tests.
    # Tcl's garbage collector must not finalize old test interpreters on workers.
    class Variable:
        def __init__(self, master=None, value=None): self.value=value
        def get(self): return self.value
        def set(self, value): self.value=value
    for name in ('StringVar', 'BooleanVar'):
        monkeypatch.setattr(ui.tk, name, Variable)
    monkeypatch.setattr(ui.tk.Toplevel, '__init__', lambda *a, **kw: None)
    for name in ('title', 'transient', 'geometry', 'minsize', 'protocol', 'bind', 'grab_set'):
        monkeypatch.setattr(ui.XmlWritebackDialog, name, lambda *a, **kw: None)
    def build(d):
        d.field_box = Control(d.field_choice)
        d.formula_check = Control(); d.summary = Control(); d.tree = Tree()
        d.review_button = Control(); d.export_button = Control(); d.close_button = Control()
        d.progress = Control(); d.controls = [(d.field_box, 'readonly')]
        d.messages = ''; d.scheduled = []
    monkeypatch.setattr(ui.XmlWritebackDialog, '_build_ui', build)
    monkeypatch.setattr(ui.XmlWritebackDialog, '_text', lambda d, s: setattr(d, 'messages', s))
    monkeypatch.setattr(ui.XmlWritebackDialog, 'after', lambda d, ms, cb: d.scheduled.append(cb))
    monkeypatch.setattr(ui.XmlWritebackDialog, 'destroy', lambda d: setattr(d, 'destroyed', True))
    strings.set_language('en')
    d = ui.XmlWritebackDialog(SimpleNamespace(winfo_toplevel=lambda: None))
    d.destroyed = False
    yield d
    if d.worker and d.worker.is_alive():
        d.cancel_event.set(); d.worker.join(timeout=10)
        assert not d.worker.is_alive()
    strings.set_language('en')


def drain(d):
    d.worker.join(timeout=10)
    assert not d.worker.is_alive()
    assert len(d.scheduled) == 1
    d.scheduled.pop()()
    assert not d.busy


def select_inputs(d, tmp_path, monkeypatch, *, amounts=None, formulas=None, declared=True):
    source = workbook(tmp_path/'มูลค่า.xlsx', **({'amounts':amounts} if amounts else {}), formulas=formulas)
    target = p6(tmp_path/'เป้าหมาย.xml', declared=declared)
    for kind, path in (('workbook', source), ('target', target)):
        monkeypatch.setattr(ui.filedialog, 'askopenfilename', lambda **kw: str(path))
        d._pick(kind)
    drain(d)
    return source, target


def choose(d, index=0):
    d.field_choice.set(d.field_box.options['values'][index]); d._invalidate()
    d.monetary.set(True); d._invalidate()


def review(d):
    d._review(); drain(d)


def export(d, out, monkeypatch):
    monkeypatch.setattr(ui.filedialog, 'asksaveasfilename', lambda **kw: str(out))
    d._export(); drain(d)


def test_discovery_never_approves_or_autoselects(dialog, tmp_path, monkeypatch):
    source, target = select_inputs(dialog, tmp_path, monkeypatch)
    assert not dialog.field_choice.get()
    assert dialog.assessment is None
    assert dialog.export_button.options['state']=='disabled'
    assert 'p6:udf:' not in dialog.field_box.options['values'][0]
    choose(dialog)
    request=dialog._request()
    assert (request.workbook,request.target_xml)==(source,target)
    assert request.field_identity=='p6:udf:47'
    assert request.monetary_confirmed and not request.recalculated_and_saved
    review(dialog)
    assert dialog.assessment.valid
    assert dialog.assessment.zero_count==1
    assert dialog.assessment.preserved_milestone_count==1
    assert dialog.assessment.target_only_ids==('NKC2-A1',)
    assert dialog.export_button.options['state']=='normal'
    assert '12.34567890123456789' in str(dialog.tree.rows)
    assert '500.50' in str(dialog.tree.rows)
    before=(source.read_bytes(),target.read_bytes())
    out=tmp_path/'ส่งออก.xml';export(dialog,out,monkeypatch)
    assert out.exists() and str(out) in dialog.status.get()
    assert before==(source.read_bytes(),target.read_bytes())
    assert dialog.assessment is None
    assert 'not a Primavera P6 import' in dialog.messages


@pytest.mark.parametrize('change',['field','money','formula'])
def test_change_invalidates_review(dialog,tmp_path,monkeypatch,change):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    if change=='field': dialog.field_choice.set(dialog.field_box.options['values'][-1])
    elif change=='money': dialog.monetary.set(False)
    else: dialog.recalculated.set(True)
    dialog._invalidate()
    assert dialog.assessment is None and not dialog.tree.rows
    assert dialog.export_button.options['state']=='disabled'
    picker=Mock();monkeypatch.setattr(ui.filedialog,'asksaveasfilename',picker)
    dialog._export();picker.assert_not_called()


def test_file_reselection_resets_confirmations_and_field(dialog,tmp_path,monkeypatch):
    source,_=select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    dialog.recalculated.set(True)
    monkeypatch.setattr(ui.filedialog,'askopenfilename',lambda **kw:str(source))
    dialog._pick('workbook');drain(dialog)
    assert not dialog.monetary.get() and not dialog.recalculated.get()
    assert not dialog.field_choice.get() and dialog.assessment is None


@pytest.mark.parametrize('picker',['open','save'])
def test_cancel_picker_does_not_export(dialog,tmp_path,monkeypatch,picker):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    original=dialog.assessment
    service=Mock(wraps=dialog.service);dialog.service=service
    monkeypatch.setattr(ui.filedialog,'askopenfilename',lambda **kw:'')
    monkeypatch.setattr(ui.filedialog,'asksaveasfilename',lambda **kw:'')
    if picker=='open':dialog._pick('target')
    else:dialog._export()
    service.export.assert_not_called()
    assert dialog.assessment is original and not dialog.busy


@pytest.mark.parametrize('confirmed',[False,True])
def test_monetary_confirmation(dialog,tmp_path,monkeypatch,confirmed):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);dialog.monetary.set(confirmed)
    review(dialog)
    assert dialog.assessment.valid is confirmed
    assert (dialog.export_button.options['state']=='normal') is confirmed


@pytest.mark.parametrize('cached', ['123.456',None,'#VALUE!'])
def test_formula_prerequisite_and_invalid_caches(dialog,tmp_path,monkeypatch,cached):
    select_inputs(dialog,tmp_path,monkeypatch,amounts=(cached,'0','0'),formulas=('100+23.456',None,None))
    assert dialog.formula_check.visible
    choose(dialog);review(dialog)
    assert not dialog.assessment.valid
    dialog.recalculated.set(True);dialog._invalidate();review(dialog)
    assert dialog.assessment.valid is (cached=='123.456')
    assert 'freshness' in dialog.messages


@pytest.mark.parametrize('declared',[False,True])
def test_explicit_new_field_and_conflict(dialog,tmp_path,monkeypatch,declared):
    source,target=select_inputs(dialog,tmp_path,monkeypatch,declared=declared)
    if declared:
        target.write_text(target.read_text(encoding='utf-8').replace('>Budget<','>Progress Studio Amount<'),encoding='utf-8')
    choose(dialog,len(dialog.fields));review(dialog)
    if declared:
        assert dialog.assessment is None  # changed field catalogue requires reselection
        choose(dialog,len(dialog.fields));review(dialog)
    assert dialog.assessment.valid is not declared
    if not declared:
        assert dialog.assessment.creates_field
        export(dialog,tmp_path/'new.xml',monkeypatch)
        assert 'Progress Studio Amount' in (tmp_path/'new.xml').read_text(encoding='utf-8')
    else:assert 'already exists' in dialog.messages


@pytest.mark.parametrize('amounts',[('0','0','0'),('1','9','0')])
def test_all_zero_and_no_change_exports(dialog,tmp_path,monkeypatch,amounts):
    _,target=select_inputs(dialog,tmp_path,monkeypatch,amounts=amounts);choose(dialog);review(dialog)
    assert dialog.assessment.valid
    if amounts[0]=='1':
        assert dialog.assessment.replacement_count==0
        assert 'No Amount replacements' in dialog.messages
    out=tmp_path/'copy.xml';export(dialog,out,monkeypatch)
    if amounts[0]=='1':assert out.read_bytes()==target.read_bytes()


@pytest.mark.parametrize('issue',['blank','negative','duplicate','milestone'])
def test_invalid_sources_never_enable_export(dialog,tmp_path,monkeypatch,issue):
    source,target=select_inputs(dialog,tmp_path,monkeypatch)
    if issue=='blank':workbook(source,amounts=(None,'0','0'))
    elif issue=='negative':workbook(source,amounts=('-1','0','0'))
    elif issue=='duplicate':workbook(source,ids=('A1','A1','M1'))
    else:workbook(source,flags=(False,False,False))
    choose(dialog);review(dialog)
    assert not dialog.assessment.valid and dialog.export_button.options['state']=='disabled'
    assert 'no partial export' in dialog.messages


@pytest.mark.parametrize('fault',['stale','collision','input','permission','parent'])
def test_output_failure_keeps_inputs_and_requires_review(dialog,tmp_path,monkeypatch,fault):
    source,target=select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    out=tmp_path/'new.xml'
    if fault=='stale':source.write_bytes(source.read_bytes()+b'changed')
    elif fault=='collision':out.write_bytes(b'occupied')
    elif fault=='input':out=target
    elif fault=='parent':out=tmp_path/'absent'/'new.xml'
    else:
        from progress_studio.services import xml_writeback_service as service
        monkeypatch.setattr(service.os,'link',Mock(side_effect=PermissionError('denied by test')))
    before=(source.read_bytes(),target.read_bytes())
    export(dialog,out,monkeypatch)
    assert dialog.assessment is None and 'failed' in dialog.status.get()
    assert before==(source.read_bytes(),target.read_bytes())
    if fault=='collision':assert out.read_bytes()==b'occupied'
    elif fault!='input':assert not out.exists()
    assert not list(tmp_path.glob('.progress-studio-xml-*'))
    assert 'Review again' in dialog.messages
    diagnostic = dialog.messages
    dialog._row_details()
    assert diagnostic in dialog.messages


def test_cancellation_repeated_clicks_and_close_while_assessing(dialog,tmp_path,monkeypatch):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog)
    entered=threading.Event();release=threading.Event();real=dialog.service.assess
    def delayed(request):
        entered.set();assert release.wait(5);return real(request)
    dialog.service.assess=Mock(side_effect=delayed)
    dialog._review();assert entered.wait(5)
    dialog._review();dialog._export();dialog._pick('target')
    assert dialog.field_box.options['state']=='disabled'
    dialog.request_close();assert not dialog.destroyed
    release.set();drain(dialog)
    assert dialog.assessment is None and 'discarded' in dialog.status.get()
    dialog.service.assess.assert_called_once()
    dialog.request_close();assert dialog.destroyed


@pytest.mark.parametrize('late',[False,True])
def test_cancellation_during_export_and_after_publication(dialog,tmp_path,monkeypatch,late):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    real=dialog.service.export;out=tmp_path/'out.xml'
    def wrapper(*args,**kw):
        if not late:dialog.cancel_event.set()
        result=real(*args,**kw)
        if late:dialog.cancel_event.set()
        return result
    dialog.service.export=wrapper
    export(dialog,out,monkeypatch)
    assert out.exists() is late
    assert ('XML exported:' in dialog.status.get()) is late
    if not late:assert 'no output' in dialog.status.get()


def test_cancel_after_render_before_publication(dialog,tmp_path,monkeypatch):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    from progress_studio.services import xml_writeback_service as service
    render=service.P6AmountDocument.render
    def cancel_render(*a,**k):
        data=render(*a,**k);dialog.cancel_event.set();return data
    monkeypatch.setattr(service.P6AmountDocument,'render',cancel_render)
    out=tmp_path/'out.xml';export(dialog,out,monkeypatch)
    assert not out.exists() and 'cancelled' in dialog.status.get()


def test_thread_start_failure_and_worker_exception(dialog,tmp_path,monkeypatch):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog)
    monkeypatch.setattr(ui.threading.Thread,'start',Mock(side_effect=RuntimeError('thread unavailable')))
    dialog._review();dialog.scheduled.pop()()
    assert not dialog.busy and dialog.assessment is None
    assert 'thread unavailable' in dialog.messages


def test_save_nested_events_cannot_export_unreviewed(dialog,tmp_path,monkeypatch):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    def save(**kw):
        dialog._invalidate();return str(tmp_path/'new.xml')
    monkeypatch.setattr(ui.filedialog,'asksaveasfilename',save)
    dialog._export();assert not dialog.busy
    assert not (tmp_path/'new.xml').exists()


def test_app_entry_independent_of_session_and_duplicate_guard(monkeypatch):
    # Deliberately no MappingStore, Project Session or BOQ attributes.
    d=SimpleNamespace(winfo_exists=lambda:True,lift=Mock())
    factory=Mock(return_value=d);monkeypatch.setattr(ui,'XmlWritebackDialog',factory)
    app=SimpleNamespace()
    ProgressStudioDesktopApp._open_xml_writeback(app)
    factory.assert_called_once_with(app)
    ProgressStudioDesktopApp._open_xml_writeback(app)
    factory.assert_called_once();d.lift.assert_called_once()
    menu=inspect.getsource(ProgressStudioDesktopApp._build_menu)
    assert "label=tr('xw.title'), command=self._open_xml_writeback" in menu


def test_application_close_waits_for_actual_writeback_outcome():
    d=SimpleNamespace(winfo_exists=lambda:True,busy=True,request_close=Mock(),lift=Mock())
    app=SimpleNamespace(_xml_writeback_dialog=d,destroy=Mock())
    ProgressStudioDesktopApp._close_application(app)
    d.request_close.assert_called_once();app.destroy.assert_not_called()


def test_field_details_and_full_row_values_preserve_diagnostics(dialog,tmp_path,monkeypatch):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    dialog._field_details();assert 'p6:udf:47' in dialog.messages
    dialog._row_details();assert '12.34567890123456789' in dialog.messages
    assert 'Warnings' in dialog.messages


def test_thai_workflow_messages(dialog,tmp_path,monkeypatch):
    strings.set_language('th')
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog);review(dialog)
    assert dialog.assessment.valid and 'ผลตรวจพร้อม' in dialog.status.get()
    out=tmp_path/'ผลลัพธ์.xml';export(dialog,out,monkeypatch)
    assert 'ส่งออก XML แล้ว' in dialog.status.get() and out.exists()


def test_workers_do_not_present_results_until_ui_poll(dialog,tmp_path,monkeypatch):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog)
    main_thread=threading.get_ident();calls=[];real=dialog.service.assess
    def assess(request):
        calls.append(threading.get_ident());return real(request)
    dialog.service.assess=assess
    show=Mock(wraps=dialog._show_assessment)
    monkeypatch.setattr(dialog,'_show_assessment',show)
    dialog._review();dialog.worker.join(timeout=5)
    assert not dialog.worker.is_alive() and dialog.busy
    show.assert_not_called()
    assert calls and calls[0]!=main_thread
    dialog.scheduled.pop()()
    show.assert_called_once();assert dialog.assessment.valid


def test_unexpected_worker_error_and_poll_wait(dialog,tmp_path,monkeypatch):
    select_inputs(dialog,tmp_path,monkeypatch);choose(dialog)
    entered=threading.Event();release=threading.Event()
    def fail(request):
        entered.set();assert release.wait(5);raise RuntimeError('unexpected diagnostic')
    dialog.service.assess=fail
    dialog._review();assert entered.wait(5)
    dialog.scheduled.pop()()  # Empty queue schedules another UI poll, no premature completion.
    assert dialog.busy and len(dialog.scheduled)==1
    release.set();drain(dialog)
    assert not dialog.busy and dialog.assessment is None
    assert 'unexpected diagnostic' in dialog.messages
