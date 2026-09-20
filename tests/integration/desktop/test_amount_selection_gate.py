from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
from progress_studio.presentation.gui.app import ProgressStudioDesktopApp
from progress_studio.domain.amount_field import AmountField
from progress_studio.presentation.cli import CommandLineInterface
from progress_studio.config import SETTINGS


def test_create_blocks_unselected_or_changed_source_before_starting_worker(tmp_path,monkeypatch):
    xml=tmp_path/'input.xml';xml.write_text('<Project/>')
    errors=[]
    monkeypatch.setattr('progress_studio.presentation.gui.app.messagebox.showerror',lambda title,text:errors.append(text))
    form=SimpleNamespace(worker=None,xml_var=Mock(get=lambda:str(xml)),weight_basis_var=Mock(get=lambda:'amount'),amount_field=None)
    ProgressStudioDesktopApp._start(form)
    assert 'Select an XML Amount field' in errors[-1]
    form.amount_field=AmountField('p6:udf:91','Contract allocation','P6','Double')
    form.amount_source_digest='stale';form._reset_amount_field=Mock()
    ProgressStudioDesktopApp._start(form)
    assert 'XML file changed' in errors[-1]
    form._reset_amount_field.assert_called_once()


def test_cli_amount_selection_is_explicit():
    cli=CommandLineInterface(SETTINGS)
    options=cli.parse(['--weight-basis','amount','--amount-field','msp:field:91'])
    assert options.weight_basis=='amount' and options.amount_field=='msp:field:91'
    assert cli.parse([]).weight_basis=='equal'
