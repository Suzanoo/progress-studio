from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from progress_studio.infrastructure.excel.live_dashboard_workbook import build_live_dashboard
from progress_studio.infrastructure.excel.live_scurve_workbook import build_live_progress_contract
from progress_studio.infrastructure.excel.rebuild_workbook_reader import RebuildWorkbookReader


def _fixture(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "main"
    headers = [
        "Row Type", "WBS", "Description", "P/A", "Activity ID",
        "Outline Level", "Plan Start", "Plan Finish", "Amount",
    ]
    for col, value in enumerate(headers, 1):
        ws.cell(4, col, value)
    dates = [
        datetime(2026, 7, 10), datetime(2026, 7, 17),
        datetime(2026, 7, 24), datetime(2026, 7, 31),
        datetime(2026, 8, 7),
    ]
    for col, dt in enumerate(dates, start=10):
        ws.cell(3, col, f"W{col-9}")
        ws.cell(4, col, dt)

    ws.append(["Project Summary","","Project","P","",0,datetime(2026,7,1),datetime(2026,8,31),1000,.20,.30,.20,.20,.10])
    ws.append(["Project Summary","","Project","A","",0,None,None,0,.10,.05,.05,None,None])
    ws.append(["S-Curve","","Plan","P","",0,None,None,None,.20,.30,.20,.20,.10])
    ws.append(["S-Curve","","Acc. Plan","AP","",0,None,None,None,.20,.50,.70,.90,1.00])
    ws.append(["S-Curve","","Actual","A","",0,None,None,None,.10,.05,.05,None,None])
    ws.append(["S-Curve","","Acc. Actual","AA","",0,None,None,None,.10,.15,.20,None,None])
    wb.save(path)
    return path


def test_lw11_progress_is_only_curve_cutoff_owner(tmp_path: Path) -> None:
    path = _fixture(tmp_path / "p.xlsx")
    dataset = RebuildWorkbookReader().read_main_dataset(path)
    wb = load_workbook(path)
    wb.create_sheet("Dashboard")["K5"] = datetime(2026, 7, 24)

    count = build_live_progress_contract(wb, dataset)
    assert count == 5
    assert wb["progress"]["A1"].value == "Date"
    assert "Dashboard!$K$5" not in wb["main"]["J10"].value
    assert "Dashboard!$K$5" in wb["progress"]["C2"].value
    assert "Dashboard!$K$5" in wb["progress"]["C6"].value
    assert "Dashboard!$K$5" not in wb["progress"]["B6"].value
    wb.close()


def test_lw11_dashboard_data_is_progress_renderer(tmp_path: Path) -> None:
    path = _fixture(tmp_path / "p.xlsx")
    dataset = RebuildWorkbookReader().read_main_dataset(path)
    wb = load_workbook(path)
    build_live_dashboard(wb, dataset, cutoff=datetime(2026, 7, 24))

    data = wb["Dashboard_Data"]
    assert data["B2"].value == "='progress'!B2"
    assert data["C2"].value == "='progress'!C2"
    # July monthly point is the final July progress row (31-Jul, progress row 5).
    assert data["E2"].value == "='progress'!B5"
    assert data["F2"].value == "='progress'!C5"
    # Dashboard selector does not duplicate cutoff behavior.
    assert "Dashboard!$K$5" not in data["I2"].value
    wb.close()
