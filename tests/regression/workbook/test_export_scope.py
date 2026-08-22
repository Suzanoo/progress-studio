from pathlib import Path

from tests._paths import REPO_ROOT

ROOT = REPO_ROOT


def test_scurve_preview_is_removed_from_current_desktop_scope() -> None:
    app_source = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    assert "S-Curve" not in app_source
    assert "SCurve" not in app_source
    assert not (ROOT / "progress_studio/presentation/gui/scurve_chart.py").exists()
    assert not (ROOT / "progress_studio/services/scurve_service.py").exists()


def test_no_hover_or_tooltip_rendering_added_to_mapping_ui() -> None:
    mapping_source = (ROOT / "progress_studio/presentation/gui/amount_mapping.py").read_text(encoding="utf-8").lower()
    assert "tooltip" not in mapping_source
    assert "<motion>" not in mapping_source
    assert "<motion>" not in mapping_source
