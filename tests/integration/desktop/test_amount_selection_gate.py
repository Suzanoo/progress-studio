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
    form=SimpleNamespace(worker=None,xml_var=Mock(get=lambda:str(xml)),weight_basis_var=Mock(get=lambda:'amount'),amount_field=None,
                         amount_initialization_var=Mock(get=lambda:'xml_field'))
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


def test_manual_creation_starts_worker_without_xml_field_preview(tmp_path, monkeypatch):
    xml=tmp_path/'input.xml';xml.write_text('<Project/>',encoding='utf-8')
    thread=Mock()
    factory=Mock(return_value=thread)
    monkeypatch.setattr('progress_studio.presentation.gui.app.threading.Thread',factory)
    errors=Mock();monkeypatch.setattr('progress_studio.presentation.gui.app.messagebox.showerror',errors)
    form=SimpleNamespace(worker=None,xml_var=Mock(get=lambda:str(xml)),
        weight_basis_var=Mock(get=lambda:'amount'),amount_initialization_var=Mock(get=lambda:'manual_placeholder'),
        amount_field=None,cutoff_var=Mock(get=lambda:'5 - Friday'),distribution_var=Mock(get=lambda:'flat'),
        progress_var=Mock(),status_var=Mock(),step_var=Mock(),run_button=Mock(),
        open_file_button=Mock(),open_folder_button=Mock(),log=Mock(),_run_worker=Mock())
    ProgressStudioDesktopApp._start(form)
    errors.assert_not_called();thread.start.assert_called_once()
    options=factory.call_args.kwargs['args'][0]
    assert options.amount_initialization=='manual_placeholder' and options.amount_field is None


def test_amount_source_controls_toggle_without_confirmation():
    form=SimpleNamespace(weight_basis_var=Mock(get=lambda:'amount'),
        amount_initialization_var=Mock(get=lambda:'manual_placeholder'),
        amount_controls=Mock(),amount_xml_controls=Mock(),amount_later_hint=Mock())
    ProgressStudioDesktopApp._weight_changed(form)
    form.amount_xml_controls.pack_forget.assert_called_once()
    form.amount_later_hint.pack.assert_called_once()
    form.amount_initialization_var.get=lambda:'xml_field'
    ProgressStudioDesktopApp._weight_changed(form)
    form.amount_xml_controls.pack.assert_called_once()
    form.amount_later_hint.pack_forget.assert_called_once()
    # All newly wired amount-source commands must name an actual method.
    import ast,inspect,textwrap
    source=ast.parse(textwrap.dedent(inspect.getsource(ProgressStudioDesktopApp._build_generator_panel)))
    for node in ast.walk(source):
        if isinstance(node,ast.keyword) and node.arg=='command' and isinstance(node.value,ast.Attribute):
            assert hasattr(ProgressStudioDesktopApp,node.value.attr)
