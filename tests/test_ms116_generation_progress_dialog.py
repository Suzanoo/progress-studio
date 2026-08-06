from __future__ import annotations

from pathlib import Path

from progress_studio.services.workbook_generation_service import WorkbookGenerationService


class _Row:
    is_summary = False


class _Source:
    project_name = "Test Project"

    def activities(self):
        return [_Row()]


def test_generation_service_exposes_optional_progress_callback() -> None:
    import inspect

    signature = inspect.signature(WorkbookGenerationService.generate)
    assert "progress_callback" in signature.parameters


def test_generation_dialog_lists_core_generation_steps() -> None:
    source = Path("progress_studio/presentation/gui/generation_progress.py").read_text(
        encoding="utf-8"
    )
    for step in (
        "Read working schedule",
        "Build main schedule",
        "Build timescale",
        "Apply amount mapping",
        "Build progress sheets",
        "Generate plan distribution",
        "Build OKD sheets",
        "Build Excel dashboard",
        "Finalize workbook",
    ):
        assert step in source


def test_export_ui_passes_generation_progress_callback() -> None:
    source = Path("progress_studio/presentation/gui/amount_mapping.py").read_text(
        encoding="utf-8"
    )
    assert "GenerationProgressDialog" in source
    assert "progress_callback=report" in source
    assert '"Generating workbook..."' in source


def test_export_generation_runs_on_background_thread_and_queues_ui_events() -> None:
    source = Path("progress_studio/presentation/gui/amount_mapping.py").read_text(
        encoding="utf-8"
    )
    assert "threading.Thread(" in source
    assert "queue.Queue" in source
    assert 'events.put(("progress"' in source
    assert "self.after(50, poll_events)" in source
    assert 'name="progress-studio-workbook-export"' in source
