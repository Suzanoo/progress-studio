import json
import tempfile
from pathlib import Path

from tests._paths import REPO_ROOT

from progress_studio.infrastructure.layout_preferences import (
    LayoutPreferences,
    LayoutPreferencesRepository,
)


ROOT = REPO_ROOT


def test_layout_preferences_round_trip():
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / "layout.json"
        repository = LayoutPreferencesRepository(path)
        expected = LayoutPreferences(True, True, 640)
        repository.save(expected)
        assert repository.load() == expected
        assert json.loads(path.read_text(encoding="utf-8"))["mapping_sash"] == 640


def test_invalid_layout_preferences_fall_back_to_defaults():
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / "layout.json"
        path.write_text("not-json", encoding="utf-8")
        assert LayoutPreferencesRepository(path).load() == LayoutPreferences()


def test_workspace_uses_lightweight_native_controls():
    amount_mapping = (ROOT / "progress_studio/presentation/gui/amount_mapping.py").read_text(encoding="utf-8")
    app = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    assert "ttk.Panedwindow" in amount_mapping
    assert "Focus Mapping" in app
    assert "Workbook Inputs" in amount_mapping
    assert "tooltip" not in amount_mapping.lower()
    assert "<Motion>" not in amount_mapping


def test_ms7_historical_documents_are_archived_and_active_docs_exist():
    assert (ROOT / "ARCHITECTURE.md").is_file()
    assert (ROOT / "ROADMAP.md").is_file()
    assert (ROOT / "docs/USER_WORKFLOW.md").is_file()
    assert (ROOT / "docs/history/milestones/MS7.md").is_file()
    assert (ROOT / "docs/history/acceptance/MS7_ACCEPTANCE.md").is_file()
    roadmap = (ROOT / "ROADMAP.md").read_text(encoding="utf-8")
    assert "P11 Production Release" in roadmap
