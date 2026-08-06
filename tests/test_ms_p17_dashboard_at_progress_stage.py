from pathlib import Path

from openpyxl import Workbook, load_workbook


def _source_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "main"
    ws.append([])
    ws.append([])
    ws.append([])
    ws.append(["Row Type", "WBS", "Description", "P/A", "Activity ID", "Outline Level"])
    ws.append(["Project Summary", "", "Demo Project", "P", "", 0])
    ws.append(["", "", "", "A", "", ""])
    wb.save(path)
    wb.close()


def test_progress_service_creates_dashboard_and_theme_before_mapping(tmp_path, monkeypatch):
    from progress_studio.services.progress_service import ProgressService
    import progress_studio.services.progress_service as module

    source = tmp_path / "source.xlsx"
    output = tmp_path / "progress.xlsx"
    _source_workbook(source)

    calls = []
    def fake_prepare(workbook, ws):
        progress = workbook.create_sheet("progress")
        progress.append(["project_start", "project_finish", "week_start", "plan", "actual"])
        progress.append([None, None, None, 0, 0])
        table = workbook.create_sheet("progress_table")
        table.append(["WBS", "Activities", "Amount", "P/A", "%Progress"])
        return (0, 0, 0, 0, 0)
    monkeypatch.setattr(module, "prepare_progress_and_scurve", fake_prepare)
    monkeypatch.setattr(module, "apply_activity_data_wbs_hierarchy", lambda ws: calls.append("theme"))
    original_dashboard = module.build_dashboard
    def wrapped_dashboard(workbook, *, project_name=None):
        calls.append("dashboard")
        return original_dashboard(workbook, project_name=project_name)
    monkeypatch.setattr(module, "build_dashboard", wrapped_dashboard)

    ProgressService().build(source, output)

    assert calls == ["theme", "dashboard"]
    result = load_workbook(output, data_only=False)
    try:
        assert result.sheetnames[0] == "Dashboard"
        assert result["Dashboard"]["C5"].value == "Demo Project"
        assert "Dashboard_Data" in result.sheetnames
    finally:
        result.close()


def test_full_generation_has_no_late_dashboard_step():
    from progress_studio.services.workbook_generation_service import WorkbookGenerationService

    service = WorkbookGenerationService()
    assert not hasattr(service, "dashboard")
