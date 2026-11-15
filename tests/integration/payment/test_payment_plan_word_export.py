from dataclasses import replace
from hashlib import sha256
from zipfile import ZipFile

import pytest
from docx import Document

from progress_studio.services import payment_plan_export as export
from tests.unit.payment.test_payment_plan_report import source


@pytest.fixture
def inputs(tmp_path):
    p = tmp_path/'source.xlsx'
    p.write_bytes(b'unit fixture for revision verification')
    values = ['.99989999','.9999',None,'.99995','.99991','.86','1',None]
    s = replace(source(values, [None]*8), revision=sha256(p.read_bytes()).hexdigest())
    return s, s.select(['W','A'],revision=s.revision), p, tmp_path


def test_native_reports_threshold_adjustments_and_unknown(inputs):
    s, selection, p, parent = inputs
    out = export.export_payment_plans(*inputs,'Plans')
    for kind in ('PERIOD','CUMULATIVE'):
        path = out/f'Plans_{kind}_PAYMENT_PLAN_REPORT.docx'
        doc = Document(path)
        assert doc.paragraphs[0].text==kind.title()+' Payment Plan Report'
        assert len(doc.tables)==9 and not doc.inline_shapes
        assert not doc.settings.element.xpath('.//w:documentProtection')
        assert all(t.rows[0]._tr.xpath('./w:trPr/w:tblHeader') for t in doc.tables)
        assert 'Date unavailable' in '\n'.join(p.text for p in doc.paragraphs)
        with ZipFile(path) as z:
            assert b'PAGE' in z.read('word/footer1.xml')
        def cells(period):
            return [[c.text for c in row.cells] for row in doc.tables[period].rows]
        assert cells(1)[1][1]=='99.99%' and 'Complete' not in cells(1)[1][0]
        assert 'Complete' in cells(2)[1][0] and 'Included in' in cells(2)[2][0]
        assert cells(2)[1][1] == ('100.00%' if kind=='CUMULATIVE' else '0.00%')
        assert 'Adjusted' in cells(4)[1][0] and 'Decrease' in cells(5)[1][0]
        assert 'Reopened' in cells(6)[1][0] and 'Completed again' in cells(7)[1][0]
        if kind=='PERIOD':
            assert cells(5)[1][2]=='-0.04' and cells(3)[-1][-1]=='0.00'
        else:
            assert cells(2)[1][2]=='999.90' and cells(3)[-2][-1]==cells(3)[-1][-1]=='999.90'
    unknown=s.select(['root'],revision=s.revision)
    folder=export.export_payment_plans(s,unknown,p,parent,'Incomplete')
    first=Document(folder/'Incomplete_CUMULATIVE_PAYMENT_PLAN_REPORT.docx').tables[1]
    assert first.rows[-1].cells[-1].text=='Unavailable'
    assert 'Incomplete' in first.rows[1].cells[0].text


@pytest.mark.parametrize('bad',['..','a/b','bad.','CON','bad:name',''])
def test_reject_invalid_folder(inputs,bad):
    with pytest.raises(ValueError):
        export.export_payment_plans(*inputs,bad)


def test_collisions_source_revision_and_empty_selection(inputs):
    s, selection, p, parent = inputs
    for name in ['Plans',p.name]:
        if name=='Plans':
            (parent/name).mkdir()
            (parent/name/'edited.docx').write_bytes(b'user edits')
        with pytest.raises(FileExistsError):
            export.export_payment_plans(*inputs,name)
    assert (parent/'Plans'/'edited.docx').read_bytes()==b'user edits'
    with pytest.raises(ValueError):
        export.export_payment_plans(s,s.select([],revision=s.revision),p,parent,'Empty')
    with pytest.raises(ValueError):
        export.export_payment_plans(replace(s,revision='other'),selection,p,parent,'Stale')
    p.write_bytes(b'different')
    with pytest.raises(ValueError,match='changed'):
        export.export_payment_plans(*inputs,'Changed')


@pytest.mark.parametrize('fault',['second','cancel','locked','modified','collision'])
def test_no_partial_publication(inputs,monkeypatch,fault):
    s, selection, p, parent = inputs
    original = export.render
    cancel = False
    def render(path,*args,**kwargs):
        nonlocal cancel
        if kwargs['cumulative']:
            if fault=='second':
                raise OSError('out of space')
            if fault=='cancel':
                cancel=True
            if fault=='modified':
                p.write_bytes(b'changed during export')
            if fault=='collision':
                (parent/'Out').mkdir()
                (parent/'Out'/'keep').write_bytes(b'existing')
        original(path,*args,**kwargs)
    monkeypatch.setattr(export,'render',render)
    if fault=='locked':
        def locked(*args):
            raise PermissionError('permission denied')
        monkeypatch.setattr(export.Path,'rename',locked)
    with pytest.raises((OSError,ValueError,export.ExportCancelled)):
        export.export_payment_plans(*inputs,'Out',cancelled=lambda:cancel)
    assert not list(parent.glob('.payment-plan-*'))
    if fault=='collision':
        assert [x.name for x in (parent/'Out').iterdir()]==['keep']
    else:
        assert not (parent/'Out').exists()
