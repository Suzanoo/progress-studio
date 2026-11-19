"""UX-1 safety contracts. Live Tk cases skip only when no display is available."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import tkinter as tk
from tkinter import ttk
import pytest
from progress_studio.presentation.gui.amount_mapping import AmountMappingFrame
from progress_studio.presentation.gui.payment import PaymentFrame
from progress_studio.presentation.gui.app import ProgressStudioDesktopApp
from progress_studio.presentation.gui.interaction import active_operations, install_native_controls
from progress_studio.domain.mapping_models import ActivityRow, BOQRow


@pytest.fixture
def root():
    try:
        window = tk.Tk()
    except tk.TclError:
        pytest.skip('Live Tk display unavailable')
    window.geometry('800x640')
    yield window
    window.destroy()
    import gc
    gc.collect()


@pytest.fixture
def mapping(root, monkeypatch):
    from progress_studio.presentation.gui.amount_mapping import LayoutPreferencesRepository
    monkeypatch.setattr(LayoutPreferencesRepository, 'save', lambda *a, **k: None)
    frame = AmountMappingFrame(root)
    frame.pack(fill='both', expand=True)
    frame.store.load_activities([ActivityRow(activity_id='A1', parent_wbs='1', child_wbs='1', description='Foundation', wbs_path=(('1', 'Civil'),))])
    frame.store.load_boq([BOQRow(key=f'B{i}', source_sheet='BOQ', source_row=i, wbs2='Civil', wbs3='Concrete', wbs4='Foundation', description=f'Item {i}', amount=100) for i in range(1, 5)])
    frame._render_activities()
    frame._render_boq()
    root.update()
    return frame


def test_payment_edit_invalidates_source_and_never_runs_old_file(root):
    frame = PaymentFrame(root, service=Mock())
    frame.workbook_var.set('old.xlsx')
    frame._validated_workbook = Path('old.xlsx')
    assert frame._execution_source() == Path('old.xlsx')
    frame.workbook_var.set('new.xlsx')
    assert frame._execution_source() is None
    assert str(frame.prepare_button.cget('state')) == 'disabled'
    assert str(frame.breakdown_button.cget('state')) == 'disabled'
    frame.service.prepare_embedded_payment_input.assert_not_called()


def test_payment_worker_error_reaches_main_thread_queue(root):
    service = Mock()
    service.prepare_embedded_payment_input.side_effect = OSError('disk full')
    frame = PaymentFrame(root, service=service)
    frame._prepare_worker(Path('source.xlsx'), Path('output.xlsx'), 2)
    callback, args = frame._events.get_nowait()
    assert callback == frame._failed
    assert str(args[0]) == 'disk full'


def test_mapping_save_failure_has_no_success_and_retains_dirty_state(mapping, monkeypatch):
    mapping.store.toggle_activity('A1')
    mapping.store.toggle_boq('B1')
    mapping.session_file = Path('project.progressstudio')
    monkeypatch.setattr(mapping, '_write_session', Mock(side_effect=OSError('locked')))
    monkeypatch.setattr('progress_studio.presentation.gui.amount_mapping.messagebox.showerror', Mock())
    mapping._notify = Mock()
    mapping._map()
    assert mapping.session_dirty
    assert mapping.store.activity_amount('A1') == 100
    assert not mapping.store.selected_boq_ids
    mapping._notify.assert_not_called()


def test_boq_native_space_range_clipboard_and_hierarchy(mapping, root):
    tree = mapping.boq_tree
    assert tree.parent('B1')
    tree.focus_force()
    tree.focus('B1')
    root.update()
    tree.event_generate('<space>')
    root.update()
    assert mapping.store.selected_boq_ids == {'B1'}
    tree.event_generate('<Shift-Down>')
    root.update()
    assert mapping.store.selected_boq_ids == {'B1', 'B2'}
    tree.event_generate('<Control-c>')
    root.update()
    copied = root.clipboard_get()
    assert 'B1\tItem 1\t100' in copied and 'B2\tItem 2\t100' in copied


def test_busy_mapping_rejects_structure_and_selection_commands(mapping):
    mapping.operation_active = True
    before = set(mapping.store.selected_boq_ids)
    mapping._select_all_filtered_boq()
    mapping._add_supplemental_activity()
    assert mapping.store.selected_boq_ids == before
    assert len(mapping.store.activities_by_id) == 1


def test_compact_mapping_has_two_trees_and_action_strip(mapping, root):
    mapping._set_inputs_collapsed(True, persist=False)
    root.update()
    assert mapping.activity_tree.winfo_height() > 100
    assert mapping.boq_tree.winfo_height() > 100
    assert mapping.mapping_status_bar.winfo_ismapped()
    assert mapping.mapping_status_bar.winfo_rooty() < root.winfo_rooty() + root.winfo_height()
    mapping.store.toggle_activity('A1')
    mapping.store.toggle_boq('B1')
    mapping._autosave_session = Mock(return_value=False)
    mapping._map()
    assert mapping.store.activity_amount('A1') == 100


def test_untracked_report_dialog_blocks_close(root, monkeypatch):
    dialog = tk.Toplevel(root)
    dialog.busy = True
    root.finance_workspace = SimpleNamespace(busy=False)
    root._save_layout_preferences = Mock()
    root.destroy = Mock(wraps=root.destroy)
    warn = Mock()
    monkeypatch.setattr('progress_studio.presentation.gui.app.messagebox.showwarning', warn)
    ProgressStudioDesktopApp._close_application(root)
    root.destroy.assert_not_called()
    warn.assert_called_once()
    assert dialog in active_operations(root)
    dialog.busy = False
    assert not active_operations(root)


def test_text_editor_keeps_undo_and_delete_ownership(root):
    entry = ttk.Entry(root)
    entry.insert(0, '123')
    root.current_workspace = 'mapping'
    root.amount_mapping = Mock()
    event = SimpleNamespace(widget=entry)
    assert ProgressStudioDesktopApp._route_command(root, event, 'undo_mapping') is None
    assert ProgressStudioDesktopApp._route_command(root, event, 'unmap_selection') is None
    root.amount_mapping.undo_mapping.assert_not_called()
    root.amount_mapping.unmap_selection.assert_not_called()
    install_native_controls(root)
    assert entry.bind('<Button-3>')


def test_payment_thread_start_failure_restores_input(root, monkeypatch):
    frame = PaymentFrame(root, service=Mock())
    frame.workbook_var.set('source.xlsx')
    frame._validated_workbook = Path('source.xlsx')
    monkeypatch.setattr('progress_studio.presentation.gui.payment.filedialog.asksaveasfilename', lambda **k: 'output.xlsx')
    monkeypatch.setattr('progress_studio.presentation.gui.payment.threading.Thread.start', Mock(side_effect=RuntimeError('cannot start')))
    monkeypatch.setattr('progress_studio.presentation.gui.payment.messagebox.showerror', Mock())
    frame._prepare()
    assert not frame._busy
    assert str(frame.source_entry.cget('state')) == 'normal'
    assert 'failed' in frame.result_var.get().lower()


def test_dirty_workspace_replacement_can_be_cancelled(mapping, monkeypatch):
    mapping.session_dirty = True
    monkeypatch.setattr('progress_studio.presentation.gui.amount_mapping.messagebox.askyesno', lambda *a, **k: False)
    assert not mapping._confirm_replace()
    mapping.set_progress_workbook(Path('replacement.xlsx'))
    assert mapping.store.activities_by_id['A1'].description == 'Foundation'
    assert mapping.session_dirty
