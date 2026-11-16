import pytest
from openpyxl import load_workbook

from progress_studio.domain.mapping_models import BOQRow
from progress_studio.infrastructure.excel.mapping_reader import ProgressActivityReader
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
from progress_studio.services.mapping_store import MappingStore
from progress_studio.services.payment_service import PaymentService
from progress_studio.services.workbook_export_service import WorkbookExportService
from progress_studio.services.xml_writeback_service import WritebackRequest, XmlWritebackService
from tests.integration.create_progress.test_amount_first_workflow import _create


def payment_sheets(path):
    workbook = load_workbook(path, read_only=True)
    try:
        return {"Payment Input", "Payment"}.intersection(workbook.sheetnames)
    finally:
        workbook.close()


@pytest.fixture
def mapping(tmp_path, monkeypatch):
    source = _create(tmp_path, monkeypatch, "p6")
    assert payment_sheets(source) == set()
    store = MappingStore()
    store.load_activities(ProgressActivityReader().read(source))
    store.load_boq([BOQRow("BOQ|2", "BOQ", 2, "1", "", "", "Concrete", 1000, "B1")])
    store.selected_activity_ids = {"A1"}
    store.selected_boq_ids = {"BOQ|2"}
    store.map_selected(75)
    return source, store


def test_mapping_and_repeated_amount_updates_do_not_initialize_payment(tmp_path, mapping):
    source, store = mapping
    for share in (75, 50):
        store.selected_boq_ids = {"BOQ|2"}
        store.map_selected(share)
        output = tmp_path / f"mapping-{share}.xlsx"
        result = WorkbookExportService().export(source, output, store)
        assert result.validation.allocated_amount == 1000 * share / 100
        assert payment_sheets(output) == set()
        source = output

    original = source.read_bytes()
    target = tmp_path / "p6.xml"
    original_xml = target.read_bytes()
    writeback = XmlWritebackService()
    assessment = writeback.assess(WritebackRequest(
        source, target, create_progress_studio_amount=True, monetary_confirmed=True,
    ))
    assert assessment.valid, assessment.errors
    assert assessment.write_total == 500
    published = writeback.export(assessment, tmp_path / "updated.xml", reviewed=True)
    assert published.replacement_count == 3
    assert target.read_bytes() == original_xml
    assert source.read_bytes() == original


def test_edited_workbook_controls_payment_reconciliation(tmp_path, mapping):
    source, store = mapping
    prepared = tmp_path / "prepared.xlsx"
    PaymentService().prepare_embedded_payment_input(source, prepared, periods=3)
    assert payment_sheets(prepared) == {"Payment Input"}
    wb = load_workbook(prepared)
    sheet = wb["Payment Input"]
    row = next(row[0].row for row in sheet if row[2].value == "A1")
    for col, value in enumerate((0, None, 0.5), 6):
        sheet.cell(row, col).value = value
    wb.save(prepared)
    wb.close()

    output = tmp_path / "reconciled.xlsx"
    WorkbookExportService().export(prepared, output, store)
    assert payment_sheets(output) == {"Payment Input", "Payment"}
    wb = load_workbook(output)
    record = next(row for row in wb["Payment Input"].values if row[2] == "A1")
    assert record[4:8] == (750, 0, None, 0.5)
    wb.close()

    # Edited-workbook migration also preserves the user's current main Amount.
    output = tmp_path / "migrated.xlsx"
    WorkbookExportService().export(source, output, store, edited_workbook=prepared)
    wb = load_workbook(output)
    record = next(row for row in wb["Payment Input"].values if row[2] == "A1")
    assert record[4:8] == (1, 0, None, 0.5)
    wb.close()

    output = tmp_path / "progress-only.xlsx"
    WorkbookExportService().export(prepared, output, store, edited_workbook=source)
    assert payment_sheets(output) == set()


def test_invalid_existing_payment_input_aborts_mapping(tmp_path, mapping):
    source, store = mapping
    wb = load_workbook(source)
    wb.create_sheet("Payment Input")["A1"] = "Unrecognized headers"
    wb.save(source)
    wb.close()
    before = source.read_bytes()
    output = tmp_path / "existing.xlsx"
    output.write_bytes(b"previous output")
    with pytest.raises(PaymentWorkbookError, match="header"):
        WorkbookExportService().export(source, output, store, overwrite=True)
    assert source.read_bytes() == before
    assert output.read_bytes() == b"previous output"
    assert not list(tmp_path.glob(".existing.*.tmp.xlsx"))
