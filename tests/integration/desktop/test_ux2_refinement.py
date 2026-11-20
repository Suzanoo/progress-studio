from unittest.mock import Mock
import gc
import tkinter as tk
from tkinter import ttk
import pytest

from progress_studio.presentation.gui.app import ProgressStudioDesktopApp
from progress_studio.presentation.gui.amount_mapping import AmountMappingFrame
from progress_studio.presentation.gui.payment import PaymentFrame
from progress_studio.presentation.gui.command_icons import icon_options
from progress_studio.presentation.gui.strings import tr, set_language
from progress_studio.domain.mapping_models import ActivityRow, BOQRow
from progress_studio.services.mapping_store import MappingStore


@pytest.fixture
def root():
    try:
        window = tk.Tk()
    except tk.TclError:
        pytest.skip('Live Tk display required')
    window.geometry('800x640')
    yield window
    window.destroy()
    gc.collect()
    set_language('en')


def rows():
    return [BOQRow('B1', 'BOQ', 1, 'Civil', 'Concrete', 'ฐานรากอาคารสำนักงานขนาดใหญ่', 'One', 100),
            BOQRow('B2', 'BOQ', 2, 'Civil', 'Concrete', 'Floor', 'Two', 200),
            BOQRow('B3', 'BOQ', 3, 'Interior', 'Concrete', 'Floor', 'Three', 300),
            BOQRow('B4', 'BOQ', 4, 'Landscape', '', '', 'Four', 400)]


@pytest.fixture
def mapping(root, monkeypatch):
    from progress_studio.presentation.gui.amount_mapping import LayoutPreferencesRepository
    monkeypatch.setattr(LayoutPreferencesRepository, 'save', lambda *a, **k: None)
    frame = AmountMappingFrame(root)
    frame.pack(fill='both', expand=True)
    frame.store.load_activities([ActivityRow('A1', '1', '1', 'Foundation')])
    frame.store.load_boq(rows())
    frame._refresh_boq_filter_values()
    frame._render_activities(); frame._render_boq()
    frame._set_inputs_collapsed(True, persist=False)
    root.update()
    return frame


def test_wbs4_filters_search_paging_and_allocations_are_independent():
    store = MappingStore(boq_page_size=1)
    store.load_boq(rows())
    store.selected_boq_ids = {'B3'}
    store.boq_wbs2 = 'Civil'; store.boq_wbs3 = 'Concrete'; store.boq_wbs4 = 'Floor'
    assert store.boq_page_data().ids == ('B2',)
    assert store.boq_wbs4_values('Civil', 'Concrete') == ('Floor', 'ฐานรากอาคารสำนักงานขนาดใหญ่')
    store.boq_query = 'missing'
    assert not store.boq_page_data().ids
    assert store.selected_boq_ids == {'B3'}
    store.load_boq(rows())
    assert store.boq_wbs4 == ''


def test_cascades_reset_only_descendants_and_preserve_checks(mapping):
    m = mapping
    m.store.selected_boq_ids = {'B3'}
    m.boq_wbs2_filter.set('Civil'); m._on_wbs2_filter()
    m.boq_wbs3_filter.set('Concrete'); m._on_wbs3_filter()
    m.boq_wbs4_filter.set('ฐานรากอาคารสำนักงานขนาดใหญ่'); m._on_wbs4_filter()
    assert m.store.boq_page_data().ids == ('B1',)
    assert m.boq_wbs4_filter.get() == rows()[0].wbs4
    m.boq_wbs2_filter.set('Landscape'); m._on_wbs2_filter()
    assert m.store.boq_wbs3 == m.store.boq_wbs4 == ''
    assert str(m.boq_wbs4_combo.cget('state')) == 'disabled'
    assert m.store.boq_page_data().ids == ('B4',)
    assert m.store.selected_boq_ids == {'B3'}


def test_clear_all_button_clears_checks_not_allocations_or_undo(mapping):
    m = mapping
    m.store.selected_activity_ids = {'A1'}; m.store.selected_boq_ids = {'B1'}
    m.store.map_selected(50)
    history = list(m.store._undo_stack)
    m.store.selected_boq_ids = {'B2', 'B3'}
    m.mapping_action_buttons[3].invoke()
    assert not m.store.selected_activity_ids and not m.store.selected_boq_ids
    assert m.store.activity_amount('A1') == 50
    assert m.store._undo_stack == history
    m.store.undo()
    assert m.store.activity_amount('A1') == 0


@pytest.mark.parametrize('width', [600, 800, 1180])
def test_action_order_and_all_buttons_inside_window(mapping, root, width):
    root.geometry(f'{width}x640'); root.update()
    buttons = mapping.mapping_action_buttons
    assert [b.cget('text') for b in buttons] == ['Map', 'Unmap', 'Undo', 'Clear all']
    assert buttons[0].cget('style') == 'Accent.TButton'
    for b in buttons:
        assert b.winfo_ismapped()
        assert b.winfo_rootx() + b.winfo_width() <= root.winfo_rootx() + root.winfo_width()
    assert mapping.boq_tree.winfo_height() > 120


def test_payment_horizontal_handlers_and_tab_reveal(root, monkeypatch):
    handlers = {name: Mock() for name in ('_prepare_boq_requirements', '_open_boq_payment', '_open_boq_word')}
    for name, handler in handlers.items():monkeypatch.setattr(PaymentFrame, name, handler)
    frame = PaymentFrame(root); frame.pack(fill='both', expand=True); root.update()
    widgets = frame.boq_workflow.winfo_children()
    assert len(widgets) == 4
    assert isinstance(widgets[1], ttk.Label)  # Existing Save in Excel guidance, not a new save operation.
    assert len({w.winfo_y() + w.winfo_height()//2 for w in widgets}) <= 2
    for widget in widgets:
        if isinstance(widget, ttk.Button):widget.invoke()
    for handler in handlers.values():handler.assert_called_once()
    root.geometry('500x640');root.update()
    widgets[-1].event_generate('<FocusIn>');root.update()
    canvas = frame.boq_workflow._strip_canvas
    assert canvas.xview()[0] > 0
    assert widgets[-1].winfo_rootx()+widgets[-1].winfo_width() <= canvas.winfo_rootx()+canvas.winfo_width()+2


def test_menu_uses_existing_guarded_handlers_and_no_toolbar(root):
    # Real native Menu with a minimal shell; business routing remains separately regression-tested.
    routes = {}
    def defer(name):
        routes[name] = Mock()
        return routes[name]
    root._defer_mapping = defer
    for name in ('_close_application','_toggle_sidebar','_toggle_focus_mapping','_show_workspace','_open_xml_writeback'):
        setattr(root, name, Mock())
    ProgressStudioDesktopApp._build_menu(root)
    menu = root.mapping_menu
    assert menu.index('end') == 2
    for index, handler in enumerate(('open_project','save_project','export_workbook')):
        menu.invoke(index);routes[handler].assert_called_once()
    assert not hasattr(ProgressStudioDesktopApp, '_build_command_bar')


def test_failed_save_can_retry_or_save_as_without_persistent_buttons(mapping, monkeypatch, tmp_path):
    m = mapping; m.session_file = tmp_path/'p.progressstudio'
    fail = Mock(side_effect=OSError('locked'))
    monkeypatch.setattr(m, '_write_session', fail)
    monkeypatch.setattr('progress_studio.presentation.gui.amount_mapping.messagebox.showerror', Mock())
    assert not m._autosave_session()
    assert m.session_dirty
    assert not any(isinstance(w, ttk.Button) for w in m.mapping_status_bar.winfo_children())
    saved = Mock(return_value=m.session_file)
    monkeypatch.setattr(m, '_write_session', saved);monkeypatch.setattr(m, '_notify', Mock())
    m.save_project();saved.assert_called_once()
    monkeypatch.setattr(m, '_save_session', Mock());m.save_project_as();m._save_session.assert_called_once()


def test_icon_reference_lifetime_and_text_fallback(root, monkeypatch):
    import progress_studio.presentation.gui.command_icons as icons
    holder = ttk.Frame(root)
    options = icon_options(holder, 'map')
    button = ttk.Button(holder, text='Map', **options)
    name = str(options['image']);del options;gc.collect()
    assert name in root.tk.call('image', 'names')
    assert button.cget('text') == 'Map'
    monkeypatch.setattr(icons, '_make_icon', Mock(side_effect=OSError('image unavailable')))
    assert icon_options(holder, 'missing') == {}


def test_thai_labels_and_long_wbs_value_are_preserved(root):
    set_language('th')
    frame = AmountMappingFrame(root)
    assert frame.boq_wbs4_filter.get() == tr('All')
    assert frame.mapping_action_buttons[-1].cget('text') == tr('Clear all')
    assert tr('ux2.clear_mappings') != tr('Clear all')


def test_existing_clear_mappings_remains_confirmed_and_undoable(mapping, monkeypatch):
    m = mapping
    m.store.selected_activity_ids = {'A1'}; m.store.selected_boq_ids = {'B1'}
    m.store.map_selected(50)
    confirm = Mock(return_value=False)
    monkeypatch.setattr('progress_studio.presentation.gui.amount_mapping.messagebox.askyesno', confirm)
    m._clear_all()
    assert m.store.activity_amount('A1') == 50
    confirm.return_value = True
    monkeypatch.setattr(m, '_autosave_session', Mock(return_value=False))
    m._clear_all()
    assert m.store.activity_amount('A1') == 0
    m.store.undo()
    assert m.store.activity_amount('A1') == 50


def test_thai_action_buttons_fit_narrow_focus_window(root):
    set_language('th')
    root.geometry('600x640')
    m = AmountMappingFrame(root); m.pack(fill='both', expand=True)
    m._set_inputs_collapsed(True, persist=False); root.update()
    for button in m.mapping_action_buttons:
        assert button.winfo_ismapped()
        assert button.winfo_rootx()+button.winfo_width() <= root.winfo_rootx()+root.winfo_width()
