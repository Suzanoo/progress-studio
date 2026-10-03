"""Headless command/routing tests; not Windows visual or Excel acceptance."""
import ast
import inspect
import json
from pathlib import Path
from string import Formatter
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
import tkinter as tk
from PIL import Image

from tests._paths import REPO_ROOT as ROOT
from progress_studio.presentation.gui import rebuild, strings, quick_guide
from progress_studio.presentation.gui.app import ProgressStudioDesktopApp
from progress_studio.infrastructure.layout_preferences import LayoutPreferences, LayoutPreferencesRepository
from progress_studio.domain.rebuild_models import RebuildMode


@pytest.fixture(autouse=True)
def reset_locale():
    strings.set_language("en")
    yield
    strings.set_language("en")


def test_resource_parity_and_named_placeholders():
    assert strings.EN.keys() == strings.TH.keys()
    def fields(value):
        return {(key, spec) for _, key, spec, _ in Formatter().parse(value) if key}
    for key in strings.EN:
        assert fields(strings.EN[key]) == fields(strings.TH[key]), key
    assert strings.tr("unknown.key", locale="th") == "unknown.key"
    assert strings.tr("home.title", locale="invalid") == strings.EN["home.title"]
    assert strings.tr("language.save_failed", locale="th", error="X.xml") .endswith("X.xml")


def test_all_literal_translation_keys_have_resources():
    for path in (ROOT / "progress_studio/presentation/gui").glob("*.py"):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "tr":
                if node.args and isinstance(node.args[0], ast.Constant):
                    assert node.args[0].value in strings.EN, (path.name, node.args[0].value)


@pytest.mark.parametrize("payload", [{}, {"language": "xx"}, {"language": None}, [], "bad"])
def test_legacy_preferences_default_eng(tmp_path, payload):
    path=tmp_path / "layout.json"
    path.write_text(json.dumps(payload))
    assert LayoutPreferencesRepository(path).load().language == "en"


def test_preference_applies_only_next_launch(tmp_path, monkeypatch):
    repository=LayoutPreferencesRepository(tmp_path / "layout.json")
    variable=SimpleNamespace(get=lambda: "THA")
    app=SimpleNamespace(language_var=variable, layout_repository=repository)
    from progress_studio.presentation.gui import app as module
    message=Mock();monkeypatch.setattr(module.messagebox, "showinfo", message)
    ProgressStudioDesktopApp._language_changed(app)
    assert repository.load().language == "th"
    assert strings.tr("Home") == "Home"
    message.assert_called_once()
    # The startup path applies the persisted language, not the selector callback.
    strings.set_language(repository.load().language)
    assert strings.tr("Home") != "Home"


def test_layout_saves_preserve_selected_language(tmp_path):
    repository=LayoutPreferencesRepository(tmp_path / "layout.json")
    repository.save(LayoutPreferences(language="th", mapping_sash=123))
    app=SimpleNamespace(layout_repository=repository, sidebar_collapsed=True)
    ProgressStudioDesktopApp._save_layout_preferences(app)
    assert repository.load().language == "th"
    from progress_studio.presentation.gui.amount_mapping import AmountMappingFrame
    frame=SimpleNamespace(layout_repository=repository, layout_preferences=LayoutPreferences(),
                          inputs_collapsed=True, body=SimpleNamespace(sashpos=lambda n: 222))
    AmountMappingFrame._save_layout_preferences(frame)
    assert repository.load().language == "th"


def test_navigation_and_home_help_routes():
    assert ProgressStudioDesktopApp.VISIBLE_WORKSPACES == (
        "home", "import", "mapping", "payment", "rebuild", "settings", "help")
    registered={row[0] for row in ProgressStudioDesktopApp.WORKSPACES}
    assert {"finance", "ai"} <= registered
    assert [row[2] for row in quick_guide.HOME_CARDS] == ["import","rebuild","mapping","payment"]
    assert 'navigate("help")' in inspect.getsource(quick_guide.build_home)
    assert 'self._show_workspace("help")' in inspect.getsource(ProgressStudioDesktopApp._build_menu)
    menu=inspect.getsource(ProgressStudioDesktopApp._build_menu)
    home=inspect.getsource(quick_guide.build_home)
    assert '"finance"' not in menu + home
    assert "Open Financial Forecast" not in home


class Control:
    def __init__(self): self.options={}; self.visible=True
    def configure(self, **kwargs): self.options.update(kwargs)
    def grid(self, **kwargs): self.visible=True
    def grid_remove(self): self.visible=False
    def pack(self, **kwargs): self.visible=True
    def pack_forget(self): self.visible=False


@pytest.fixture
def frame(monkeypatch):
    interpreter=tk.Tcl()
    variable=tk.StringVar
    monkeypatch.setattr(rebuild.tk,"StringVar",lambda **kw: variable(master=interpreter, **kw))
    monkeypatch.setattr(rebuild.ttk.Frame,"__init__",lambda *a,**kw: None)
    def build(self):
        for key in ("output_options","ev_options","prepare_button","rebuild_button","open_button"):
            setattr(self,key,Control())
        self.ev_generate_button=self.rebuild_button
        self._control_states=[]
    monkeypatch.setattr(rebuild.RebuildFrame,"_build_ui",build)
    engine=Mock()
    engine.analyze.return_value=SimpleNamespace(payment_input_present=True,activity_count=3,
        existing_generated_sheets=(),contract=SimpleNamespace(generated_progress=("progress",)))
    ev=Mock()
    ev.analyze.return_value=SimpleNamespace(existing_earned_value_sheet=False,activity_count=3,project_bac=120,allocation_count=2)
    result=rebuild.RebuildFrame(None,engine,ev)
    monkeypatch.setattr(result,"after",lambda delay,callback: callback())
    return result


@pytest.mark.parametrize("locale",["en","th"])
@pytest.mark.parametrize("target",["progress","payment","ev"])
def test_conditional_options_and_canonical_sources(frame, locale, target):
    strings.set_language(locale)
    frame.workbook_var.set("source.xlsx")
    frame.target_var.set(target)
    frame._target_changed()
    assert frame.ev_options.visible == (target == "ev")
    assert frame.output_options.visible == (target != "ev")
    if target == "ev":
        frame.ev_service.analyze.assert_called_once_with(Path("source.xlsx"),monetary_source="BOQ Mapping")
        frame.engine.analyze.assert_not_called()
        assert not frame.prepare_button.visible
        frame.ev_source_var.set("Activity Amount");frame._target_changed()
        assert frame.prepare_button.visible
        assert frame.ev_service.analyze.call_args.kwargs["monetary_source"] == "Activity Amount"
    else:
        frame.engine.analyze.assert_called_once_with(Path("source.xlsx"),RebuildMode(target))
        frame.ev_service.analyze.assert_not_called()


@pytest.mark.parametrize("target",["progress","payment","ev"])
def test_single_action_routes_selected_target(frame,target):
    frame.target_var.set(target)
    frame._rebuild=Mock();frame._generate_ev=Mock()
    frame._run_target()
    assert frame._generate_ev.call_count == (target=="ev")
    assert frame._rebuild.call_count == (target!="ev")


@pytest.mark.parametrize("mode",[RebuildMode.PROGRESS,RebuildMode.PAYMENT])
@pytest.mark.parametrize("output_mode",["snapshot","live"])
def test_worker_preserves_four_engine_routes(frame,mode,output_mode):
    expected={(RebuildMode.PROGRESS,"snapshot"):"rebuild_progress",
              (RebuildMode.PROGRESS,"live"):"rebuild_live_progress",
              (RebuildMode.PAYMENT,"snapshot"):"rebuild_payment",
              (RebuildMode.PAYMENT,"live"):"rebuild_live_payment"}[mode,output_mode]
    getattr(frame.engine,expected).return_value=SimpleNamespace(activity_count=3,week_count=4,monthly_periods=1,rendered_periods=1,rendered_points=2)
    frame._worker_rebuild(Path("in.xlsx"),Path("out.xlsx"),mode,output_mode)
    assert getattr(frame.engine,expected).call_count==1
    assert frame._output_path==Path("out.xlsx")


def test_target_switch_invalidates_prior_result_and_no_ev_fallback(frame):
    frame.workbook_var.set("source.xlsx");frame._target_changed()
    assert frame._validated_path
    frame.target_var.set("ev")
    frame.ev_service.analyze.side_effect=rebuild.EarnedValueRebuildError("invalid source")
    frame._target_changed()
    assert frame._validated_path is None and frame._ev_validated_path is None
    assert frame.rebuild_button.options["state"]=="disabled"
    assert frame.engine.analyze.call_count==1


def test_icon_family_and_guide_protection_note():
    brand=ROOT / "progress_studio/assets/brand"
    assert "S-curve" in (brand / "progress_studio.svg").read_text(encoding="utf-8")
    with Image.open(brand / "progress_studio.ico") as icon:
        assert {(16,16),(32,32),(256,256)} <= icon.info["sizes"]
    from openpyxl import Workbook
    from progress_studio.infrastructure.excel.workbook_guide import build_workbook_guide
    workbook=Workbook();build_workbook_guide(workbook)
    text=" ".join(str(c.value or "") for row in workbook["README"] for c in row)
    assert "okmd" in text and "does not encrypt" in text and "may restore" in text


def test_active_document_links_exist():
    import re
    for name in ("README.md","docs/README.md","docs/ROADMAP.md","docs/ARCHITECTURE.md",
                 "docs/EV_LIVE_CONTRACT.md","docs/RELEASE_CHECKLIST.md","docs/V1_1_PRODUCT_EXPERIENCE.md"):
        path=ROOT/name
        for link in re.findall(r'\]\(([^\s)]+)\)',path.read_text(encoding="utf-8")):
            if '://' not in link and not link.startswith('#'):
                assert (path.parent / link.split('#')[0]).exists(), (name,link)
