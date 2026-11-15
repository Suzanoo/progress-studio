import json
from unittest.mock import patch

import pytest

from progress_studio.presentation.gui.payment_plan_selection import PlanSelection
from tests.unit.payment.test_payment_plan_report import source


def model():
    return PlanSelection(source(['.9999',None,'.85','1'],['0','.25',None,None]))


def test_filter_selection_overlap_and_both_previews():
    m=model()
    m.select(['A','W'])
    assert [r.key for r in m.matching('pour')]==['root','W','A']
    assert [r.key for r in m.result.rows]==['W','A']
    assert m.matching('missing')==() and m.result.included_in=={'A':'W'}
    assert m.result.selected.activities==('A',)
    v=m.result.row_values['A'].values
    assert v[0].period_visible and v[0].cumulative_visible
    assert not v[1].period_visible and not v[1].cumulative_visible
    assert m.result.retained_money[1]==m.result.selected.values[1].money
    assert v[2].correction and v[2].movement<0 and v[3].transition=='recompleted'


def test_saved_selections_are_atomic_revision_bound_and_ordered(tmp_path):
    m=model()
    m.select(['A','W'])
    path=tmp_path/'review.paymentselection.json'
    m.save(path)
    original=path.read_bytes()
    assert json.loads(original)['rows']==['W','A']
    m.select(['B'])
    with patch('progress_studio.presentation.gui.payment_plan_selection.os.replace',side_effect=OSError):
        with pytest.raises(OSError):
            m.save(path)
    assert path.read_bytes()==original and list(tmp_path.iterdir())==[path]
    m.restore(path)
    assert m.keys=={'W','A'}
    for change in [{'project':'other'},{'revision':'different'},{'version':True},
                   {'rows':['gone']},{'rows':['B','B']},{'rows':None}]:
        path.write_text(json.dumps(dict(json.loads(original),**change)),encoding='utf-8')
        with pytest.raises(ValueError):
            m.restore(path)
        assert m.keys=={'W','A'}
    with pytest.raises(ValueError):
        m.save(tmp_path/'payment.xlsx')


def test_dialog_installs_through_payment_and_preserves_previous_source(tmp_path,monkeypatch):
    import tkinter as tk
    from time import monotonic,sleep
    from progress_studio.presentation.gui import payment_plan_dialog as ui
    from progress_studio.presentation.gui.payment import PaymentFrame
    try:
        root=tk.Tk()
    except tk.TclError:
        pytest.skip('Tk display unavailable')
    try:
        frame=PaymentFrame(root)
        frame._open_plan_report()
        d=frame._payment_plan_dialog
        frame._open_plan_report()
        assert frame._payment_plan_dialog is d
        monkeypatch.setattr(ui,'read_payment_plan_source',lambda p:model().source)
        def load(name):
            d.load(tmp_path/name)
            limit=monotonic()+5
            while d.busy and monotonic()<limit:
                root.update()
                sleep(.01)
            assert not d.busy
        load('source.xlsx')
        d.apply({'W','A'})
        d.period.current(1)
        d.mode.set('cumulative')
        d.draw_preview()
        assert '999.90' in d.retained.get() and not d.preview.exists('A')
        d.period.current(2)
        d.draw_preview()
        assert d.preview.exists('A')
        d.query.set('Finish')
        d.draw_tree()
        assert not d.tree.exists('A') and d.model.keys=={'W','A'}
        previous=d.model
        monkeypatch.setattr(ui,'read_payment_plan_source',lambda p: (_ for _ in ()).throw(OSError('locked')))
        monkeypatch.setattr(ui.messagebox,'showerror',lambda *a,**kw:None)
        load('invalid.xlsx')
        assert d.model is previous and d.preview.exists('A') and 'invalid.xlsx' in d.status.get()
    finally:
        root.destroy()


def test_real_tree_mouse_keyboard_search_and_pending_close(monkeypatch):
    import tkinter as tk
    from pathlib import Path
    from threading import Event
    from time import sleep,monotonic
    from progress_studio.presentation.gui import payment_plan_dialog as ui
    from progress_studio.presentation.gui.strings import set_language
    try:
        root=tk.Tk()
    except tk.TclError:
        pytest.skip('Tk display unavailable')
    try:
        set_language('th')
        d=ui.PaymentPlanDialog(root)
        d.model=model()
        d.period.configure(values=['P01','P02','P03','P04'])
        d.period.current(0)
        d.expanded={'root','W'}
        d.draw_tree()
        d.refresh()
        root.update()
        x,y,w,h=d.tree.bbox('A','#0')
        d.tree.event_generate('<Button-1>',x=x+w//2,y=y+h//2)
        root.update()
        assert d.model.keys=={'A'}
        d.tree.focus_force()
        d.tree.event_generate('<Key-space>')
        root.update()
        assert not d.model.keys
        d.tree.item('W',open=False)
        d.query.set('Pour')
        d.draw_tree()
        assert d.tree.item('W','open')
        d.query.set('')
        d.draw_tree()
        assert not d.tree.item('W','open')
        d.apply({'W','A'})
        d.review_rows()
        win=d.grab_current()
        assert win.title()=='ตรวจรายการที่เลือก'
        win.destroy()
        d.details()
        root.update()
        assert 'บางส่วน' in d.scope.get()
        lock=Event()
        def wait_source(path):
            lock.wait(2)
            return model().source
        monkeypatch.setattr(ui,'read_payment_plan_source',wait_source)
        d.load(Path('pending.xlsx'))
        d.destroy()
        lock.set()
        root.update()
    finally:
        set_language('en')
        root.destroy()


def test_changed_revision_cancel_and_unknown_preview(tmp_path,monkeypatch):
    import tkinter as tk
    from time import monotonic,sleep
    from dataclasses import replace
    from progress_studio.presentation.gui import payment_plan_dialog as ui
    try:
        root=tk.Tk()
    except tk.TclError:
        pytest.skip('Tk display unavailable')
    try:
        d=ui.PaymentPlanDialog(root)
        old=model()
        old.select(['W'])
        d.model=old
        d.period.configure(values=['P01','P02','P03','P04'])
        d.period.current(2)
        new=replace(source(['1'],[None]),revision='different')
        monkeypatch.setattr(ui,'read_payment_plan_source',lambda p:new)
        def wait():
            stop=monotonic()+5
            while d.busy and monotonic()<stop:
                root.update()
                sleep(.01)
            assert not d.busy
        monkeypatch.setattr(ui.messagebox,'askyesno',lambda *a,**kw:False)
        d.load(tmp_path/'revision.xlsx')
        wait()
        assert d.model is old and d.model.keys=={'W'} and d.period.current()==2
        monkeypatch.setattr(ui.messagebox,'askyesno',lambda *a,**kw:True)
        d.load(tmp_path/'revision.xlsx')
        wait()
        assert d.model.source is new and not d.model.keys
        assert any('B:' in msg for msg in d.issues())
        d.apply({'root'})
        assert 'Unavailable' in d.total.get()
        assert d.preview.set('root','state')=='Incomplete'
    finally:
        root.destroy()
