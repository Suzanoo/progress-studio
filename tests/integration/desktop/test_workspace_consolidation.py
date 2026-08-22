from __future__ import annotations

from pathlib import Path

from tests._paths import REPO_ROOT

from openpyxl import Workbook

from progress_studio.infrastructure.excel.final_workbook_policy import finalize_workbook


ROOT = REPO_ROOT


def test_finish1_support_sheets_follow_proven_visibility_contract() -> None:
    wb = Workbook()
    wb.active.title = "main"
    for name in ("main_monthly", "Dashboard", "progress", "progress_table", "Dashboard_Data", "Info", "Timescale Info", "Amount Mapping", "Distribution Report"):
        wb.create_sheet(name)

    result = finalize_workbook(wb, mode="snapshot", include_guide=True)

    assert wb["main"].sheet_state == "visible"
    assert wb["main_monthly"].sheet_state == "visible"
    assert wb["Dashboard"].sheet_state == "visible"
    for name in ("progress", "progress_table", "Dashboard_Data"):
        assert wb[name].sheet_state == "hidden"
    for name in ("Info", "Timescale Info", "Amount Mapping", "Distribution Report"):
        assert wb[name].sheet_state == "veryHidden"
        assert name in result.very_hidden_sheets


def test_finish1_welcome_is_workflow_first_and_create_is_clean() -> None:
    source = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")

    assert 'self.current_workspace = "home"' in source
    assert 'self._show_workspace("home")' in source
    assert "MSP / P6 XML → Progress Workbook" in source
    assert "BOQ → Activity Amount" in source
    assert "Prepare and render payment stages" in source
    assert "Edited Workbook → Updated Workbook" in source
    assert "Mapping and Payment are optional" in source

    create_start = source.index("    def _build_import_workspace")
    create_end = source.index("    def _build_mapping_workspace", create_start)
    create_block = source[create_start:create_end]
    assert "Export Mapped Workbook" not in create_block
    assert 'text="Go to Mapping"' in create_block


def test_finish1_create_file_picker_uses_schedule_xml_wording() -> None:
    source = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    assert 'title="Select Schedule XML"' in source
    assert '("Schedule XML", "*.xml")' in source
