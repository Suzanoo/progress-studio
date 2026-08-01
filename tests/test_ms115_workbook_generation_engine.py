from pathlib import Path

from openpyxl import Workbook

from progress_studio.domain.working_tree import WorkingNodeKind, WorkingNodeOrigin, WorkingTreeNode
from progress_studio.services.working_tree_schedule_source import WorkingTreeScheduleSource


def _source(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "main"
    ws.append(["Row Type", "WBS", "Description", "P/A", "Activity ID", "Task ID", "UID", "Outline Level", "Plan Start", "Plan Finish", "Actual Start", "Actual Finish", "% Complete", "Physical %", "Amount", "Total Float (hr)"])
    ws.append(["Project Summary", "", "Demo", "P", "", None, None, 0, None, None])
    ws.append(["WBS", "1", "Structure", "P", "", None, None, 1, None, None])
    ws.append(["Activity", "1.1", "Existing", "P", "A1000", 10, 20, 2, "2026-01-01", "2026-01-05", None, None, 0.25, 0.2, 0, 8])
    wb.save(path)
    return path


def test_working_tree_source_preserves_existing_dates_and_allows_draft_dates(tmp_path: Path) -> None:
    source = _source(tmp_path / "source.xlsx")
    nodes = [
        WorkingTreeNode("w1", WorkingNodeKind.WBS, None, "1", "Structure", WorkingNodeOrigin.WORKBOOK, 0, source_path=(("1", "Structure"),)),
        WorkingTreeNode("a1", WorkingNodeKind.ACTIVITY, "w1", "A1000", "Existing", WorkingNodeOrigin.WORKBOOK, 0, source_activity_id="A1000"),
        WorkingTreeNode("a2", WorkingNodeKind.ACTIVITY, "w1", "A1010", "Draft", WorkingNodeOrigin.USER_CREATED, 1, source_activity_id="A1010"),
    ]
    rows = WorkingTreeScheduleSource(source, nodes, {"A1000": 10, "A1010": 20}).activities()
    by_id = {row.activity_id: row for row in rows if not row.is_summary}
    assert by_id["A1000"].plan_start is not None
    assert by_id["A1000"].plan_finish is not None
    assert by_id["A1010"].plan_start is None
    assert by_id["A1010"].plan_finish is None
    assert by_id["A1010"].amount == 20
