from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from progress_studio.domain.mapping_models import ActivityRow, BOQRow
from progress_studio.infrastructure.excel.mapping_reader import BOQSheetReader
from progress_studio.services.mapping_store import MappingStore
from progress_studio.services.workbook_export_service import WorkbookExportService


def make_progress(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Amount Mapping"
    ws.append(["Activity ID", "WBS", "Description", "Amount", "Status"])
    ws.append([None, "1", "Project", None, None])
    ws.append(["A1000", "1.1", "First activity", 99, "OLD"])
    ws.append(["A2000", "1.2", "Second activity", 99, "OLD"])
    formula = wb.create_sheet("Formula Check")
    formula["A1"] = "=SUM('Amount Mapping'!D3:D4)"
    wb.save(path)
    return path


def make_store() -> MappingStore:
    store = MappingStore()
    store.load_activities([
        ActivityRow("A1000", "1", "1.1", "First activity"),
        ActivityRow("A2000", "1", "1.2", "Second activity"),
    ])
    store.load_boq([
        BOQRow("Sheet|10|2", "Sheet", 10, "W2", "W3", "W4", "Item one", 1000, "BOQ-ONE"),
        BOQRow("Sheet|11|3", "Sheet", 11, "W2", "W3", "W4", "Item two", 500, "BOQ-TWO"),
    ])
    store.selected_activity_ids = {"A1000"}
    store.selected_boq_ids = {"Sheet|10|2"}
    store.map_selected(40)
    store.selected_activity_ids = {"A2000"}
    store.selected_boq_ids = {"Sheet|10|2", "Sheet|11|3"}
    store.map_selected(60)
    return store


def test_export_updates_amount_mapping_and_writes_reconciliation(tmp_path: Path) -> None:
    source = make_progress(tmp_path / "progress.xlsx")
    output = tmp_path / "progress_mapped.xlsx"
    result = WorkbookExportService().export(source, output, make_store())

    assert result.amount_rows_updated == 2
    assert result.mapping_rows_written == 3
    assert result.validation.partial_boq_count == 1
    assert result.validation.full_boq_count == 1
    assert result.validation.allocated_amount == 1300

    wb = load_workbook(output, data_only=False)
    try:
        amount = wb["Amount Mapping"]
        assert amount["D3"].value == 400
        assert amount["D4"].value == 900
        assert amount["E3"].value == "MAPPED"
        assert wb["Formula Check"]["A1"].value == "=SUM('Amount Mapping'!D3:D4)"

        mapping = wb["BOQ Activity Mapping"]
        assert mapping.max_row == 4
        assert mapping["M2"].value in {"BOQ-ONE", "BOQ-TWO"}
        assert mapping.tables

        summary = wb["Mapping Summary"]
        assert summary["B5"].value == "Partial"
        assert summary["B14"].value == 1300
        assert wb.calculation.fullCalcOnLoad is True
    finally:
        wb.close()


def test_export_refuses_source_overwrite_and_existing_output_by_default(tmp_path: Path) -> None:
    source = make_progress(tmp_path / "progress.xlsx")
    service = WorkbookExportService()
    store = make_store()
    with pytest.raises(ValueError, match="different"):
        service.export(source, source, store, overwrite=True)

    output = tmp_path / "existing.xlsx"
    output.write_bytes(b"keep")
    with pytest.raises(FileExistsError):
        service.export(source, output, store)
    assert output.read_bytes() == b"keep"
    assert not list(tmp_path.glob("*.tmp.xlsx"))


def test_validation_reports_complete_mapping() -> None:
    store = make_store()
    store.selected_activity_ids = {"A1000"}
    store.selected_boq_ids = {"Sheet|11|3"}
    store.map_selected(40)
    validation = WorkbookExportService.validate(store)
    assert validation.is_complete
    assert validation.remaining_amount == 0
    assert validation.allocated_percent == 100


def test_boq_reader_creates_repeatable_stable_export_id(tmp_path: Path) -> None:
    path = tmp_path / "boq.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "BOQ"
    ws.append(["WBS-2", "WBS-3", "WBS-4", "Description", "Amount", "Source Sheet", "Source Row"])
    ws.append(["A", "B", "C", "Concrete", 100, "Original", 55])
    wb.save(path)

    first = BOQSheetReader().read(path, "BOQ")[0]
    second = BOQSheetReader().read(path, "BOQ")[0]
    assert first.stable_id.startswith("BOQ-")
    assert first.stable_id == second.stable_id
