from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_lw0_sidebar_removes_standalone_export_workspace() -> None:
    source = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")

    assert '("export", "⇧", "Export")' not in source
    assert 'label="Export Workspace"' not in source
    assert 'def _build_export_workspace' not in source
    assert '("rebuild", "↻", "Rebuild")' in source


def test_lw0_initial_export_lives_in_create_mapping_flow() -> None:
    source = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")

    start = source.index("    def _build_import_workspace")
    end = source.index("    def _build_mapping_workspace", start)
    create_block = source[start:end]

    assert "Export Mapped Workbook" in create_block
    assert 'self._defer_mapping("export_workbook")' in create_block

    command_start = source.index("    def _build_command_bar")
    command_end = source.index("    def _new_workspace", command_start)
    command_block = source[command_start:command_end]
    assert "Export Mapped Workbook" in command_block
    assert 'self._defer_mapping("export_workbook")' in command_block


def test_lw0_rebuild_ui_separates_output_mode_from_scope() -> None:
    source = (ROOT / "progress_studio/presentation/gui/rebuild.py").read_text(encoding="utf-8")

    assert "Output Mode" in source
    assert "Snapshot Workbook" in source
    assert "Live Workbook" in source
    assert "Rebuild Scope" in source
    assert 'text="Progress"' in source
    assert 'text="Payment"' in source
    assert "output_mode_var" in source
    assert "mode_var" in source


def test_lw0_live_mode_cannot_route_into_snapshot_engine() -> None:
    source = (ROOT / "progress_studio/presentation/gui/rebuild.py").read_text(encoding="utf-8")

    assert 'self.output_mode_var.get() != "snapshot"' in source
    assert "LW-0 defines the Live Workbook UX only" in source
    assert 'self.rebuild_button.configure(state="disabled")' in source
    assert "rebuild_progress" in source
    assert "rebuild_payment" in source
