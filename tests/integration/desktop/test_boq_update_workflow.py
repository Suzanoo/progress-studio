"""Headless workflow callbacks; does not claim native desktop acceptance."""
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from tests.fixtures.boq_revision import revision_case
from progress_studio.infrastructure.session.mapping_session_repository import materialize_workbook_snapshot
from progress_studio.presentation.gui import amount_mapping as ui


@pytest.mark.parametrize('action',['confirm','cancel','failure','stale','refresh_failure'])
def test_review_update_transaction(tmp_path,monkeypatch,action):
    c=revision_case(tmp_path);old=c.project.read_bytes();old_store=c.store
    monkeypatch.setattr(c.repo,'materialize_snapshot',lambda s:materialize_workbook_snapshot(s,tmp_path/'cache'))
    info=Mock();error=Mock()
    monkeypatch.setattr(ui.messagebox,'showinfo',info)
    monkeypatch.setattr(ui.messagebox,'showerror',error)
    monkeypatch.setattr(ui.filedialog,'askopenfilename',lambda **k:str(c.new))
    def review(parent,preview):
        assert c.project.read_bytes()==old and parent.store is old_store
        if action=='stale':c.new.write_bytes(b'changed after review')
        return action!='cancel'
    monkeypatch.setattr(ui,'review_boq_update',review)
    if action=='failure':monkeypatch.setattr(c.repo,'save',Mock(side_effect=OSError('disk full')))
    frame=SimpleNamespace(session_file=c.project,progress_file=c.progress,boq_file=c.old,
                          boq_sheet='Project',session_repository=c.repo,store=c.store,
                          _session_data=lambda:c.session,boq_path_var=Mock(),session_status_var=Mock(),
                          _busy=Mock(),_refresh_boq_filter_values=Mock(),_render_boq=Mock(),
                          _render_activities=Mock(),_update_summary=Mock(),_update_input_status=Mock())
    if action=='refresh_failure':frame._render_boq.side_effect=RuntimeError('view failed')
    ui.AmountMappingFrame._update_boq(frame)
    if action in ('confirm','refresh_failure'):
        assert frame.store is not old_store
        assert frame.store.allocations==old_store.allocations
        assert c.repo.load(c.project).boq_snapshot==frame._boq_snapshot
        assert frame.boq_file.exists()
        if action=='refresh_failure':assert 'was saved' in error.call_args.args[1]
        else:error.assert_not_called()
    else:
        assert frame.store is old_store and c.project.read_bytes()==old
        if action=='cancel':error.assert_not_called()
        else:error.assert_called_once()
    frame._busy.assert_called_with(False)


def test_cancel_file_picker_does_not_build_or_save_project(monkeypatch):
    monkeypatch.setattr(ui.messagebox,'showinfo',Mock())
    monkeypatch.setattr(ui.filedialog,'askopenfilename',lambda **k:'')
    frame=SimpleNamespace(session_file='a',progress_file='b',boq_file='c',boq_sheet='Project',
                          _session_data=Mock(side_effect=AssertionError('must not run')))
    ui.AmountMappingFrame._update_boq(frame)
    frame._session_data.assert_not_called()
