from datetime import date

from openpyxl import Workbook

from progress_studio.infrastructure.excel.dashboard_workbook import (
    DASHBOARD_SHEET,
    DATA_SHEET,
    build_dashboard,
)


def _workbook():
    wb = Workbook()
    progress = wb.active
    progress.title = "progress"
    progress.append(["project_start", "project_finish", "week_start", "plan", "actual"])
    progress.append([date(2026, 1, 1), date(2026, 2, 1), date(2026, 1, 2), 10, 5])
    progress.append([date(2026, 1, 1), date(2026, 2, 1), date(2026, 1, 9), 25, 12])
    table = wb.create_sheet("progress_table")
    table.append(["WBS", "Activities", "Amount", "P/A", "%Progress", date(2026, 1, 2), date(2026, 1, 9)])
    table.append(["1.1", "Activity A", 1000, "P", 25, 10, 15])
    table.append(["1.1", "Activity A", 1000, "A", 12, 5, 7])
    return wb


def test_dashboard_is_created_as_first_separate_sheet_with_controls_and_chart():
    wb = _workbook()
    build_dashboard(wb, project_name="Demo Project")

    assert wb.sheetnames[0] == DASHBOARD_SHEET
    assert DATA_SHEET in wb.sheetnames
    assert wb[DATA_SHEET].sheet_state == "hidden"
    dashboard = wb[DASHBOARD_SHEET]
    assert dashboard["B2"].value == "PROGRESS STUDIO DASHBOARD"
    assert dashboard["C5"].value == "Demo Project"
    assert dashboard["D5"].value == "Weekly"
    assert dashboard["H5"].value.startswith("='Dashboard_Data'!")
    assert len(dashboard.data_validations.dataValidation) == 2
    assert len(dashboard._charts) == 1


def test_dashboard_activity_rows_are_formula_linked_to_progress_table():
    wb = _workbook()
    build_dashboard(wb)
    dashboard = wb[DASHBOARD_SHEET]

    assert dashboard["B38"].value == "='progress_table'!A2"
    assert dashboard["C38"].value == "='progress_table'!B2"
    assert dashboard["F38"].value == "Plan"
    assert dashboard["F39"].value == "Actual"
    assert "SUMPRODUCT" in dashboard["K38"].value


def test_progress_service_creates_dashboard_before_mapping(tmp_path):
    from progress_studio.services.progress_service import ProgressService

    source = tmp_path / "source.xlsx"
    output = tmp_path / "progress.xlsx"
    wb = _workbook()
    # ProgressService expects a main sheet and then builds progress sheets from it;
    # use a focused monkeypatch to verify orchestration without duplicating the
    # complete schedule fixture.
    main = wb.create_sheet("main", 0)
    main.append([])
    main.append([])
    main.append([])
    main.append(["Row Type", "WBS", "Description", "P/A", "Activity ID", "Outline Level"])
    main.append(["Project Summary", "", "Demo Project", "P", "", 0])
    main.append(["", "", "", "A", "", ""])
    wb.save(source)
    wb.close()

    import progress_studio.services.progress_service as module
    original = module.prepare_progress_and_scurve
    module.prepare_progress_and_scurve = lambda workbook, ws: (0, 0, 0, 0, 0)
    try:
        ProgressService().build(source, output)
    finally:
        module.prepare_progress_and_scurve = original

    from openpyxl import load_workbook
    result = load_workbook(output, data_only=False)
    try:
        assert result.sheetnames[0] == "Dashboard"
        assert result["Dashboard"]["C5"].value == "Demo Project"
        assert "Dashboard_Data" in result.sheetnames
    finally:
        result.close()
