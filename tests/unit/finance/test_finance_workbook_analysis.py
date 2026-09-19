from datetime import date

import pytest
from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904

from progress_studio.services.financial_forecast_workbook_service import FinancialForecastWorkbookService
from tests.unit.finance.test_finance_workbook import sample


def test_analysis_reads_existing_saved_settings_without_mutation(tmp_path):
    path = tmp_path / "finance.xlsx"
    wb = sample()
    wb["Finance Input"]["B5"] = date(2026, 1, 10)
    wb.save(path)
    wb.close()
    before = path.read_bytes()
    result = FinancialForecastWorkbookService().analyze(path)
    assert path.read_bytes() == before
    assert result.workbook == path
    assert result.existing_finance and result.capacity_rows == 4
    assert result.opening_date == date(2026, 1, 10)
    assert result.actuals_through == date(2026, 1, 31)
    assert result.currency == "THB"


def test_new_workbook_does_not_infer_finance_dates_or_currency(tmp_path):
    path = tmp_path / "progress.xlsx"
    wb = Workbook()
    wb.active.title = "main"
    wb.save(path)
    wb.close()
    result = FinancialForecastWorkbookService().analyze(path)
    assert not result.existing_finance
    assert result.capacity_rows == 100
    assert result.opening_date is None and result.actuals_through is None and result.currency is None


@pytest.mark.parametrize("invalid,match", [
    ("no_main", "containing 'main'"), ("1904", "1900 date system"),
    ("schema", "schema"), ("invalid_cash", "Negative"),
    ("collision", "collision"), ("header", "headers"),
])
def test_analysis_rejects_unready_saved_workbooks(tmp_path, invalid, match):
    path = tmp_path / "invalid.xlsx"
    wb = sample()
    if invalid == "no_main":
        wb.remove(wb["main"])
    elif invalid == "1904":
        wb.epoch = CALENDAR_MAC_1904
    elif invalid == "schema":
        wb["Finance Input"]["A1"] = "unknown"
    elif invalid == "invalid_cash":
        wb["Finance Input"]["C25"] = -40
    elif invalid == "collision":
        wb.create_sheet("Cash Flow")["A1"] = "User-owned cash sheet"
    elif invalid == "header":
        wb["Finance Input"]["A24"] = "Changed"
    wb.save(path)
    wb.close()
    before = path.read_bytes()
    with pytest.raises(ValueError, match=match):
        FinancialForecastWorkbookService().analyze(path)
    assert path.read_bytes() == before


@pytest.mark.parametrize("filename", ["input.xlsm", "input.xls", "input.txt"])
def test_analysis_rejects_unsupported_extension(tmp_path, filename):
    with pytest.raises(ValueError, match=".xlsx"):
        FinancialForecastWorkbookService().analyze(tmp_path / filename)


def test_analysis_reports_missing_and_corrupt_file(tmp_path):
    path = tmp_path / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        FinancialForecastWorkbookService().analyze(path)
    path.write_text("not a workbook")
    from zipfile import BadZipFile
    with pytest.raises(BadZipFile):
        FinancialForecastWorkbookService().analyze(path)


def test_analysis_accepts_incomplete_finance_for_editing(tmp_path):
    path = tmp_path / "incomplete.xlsx"
    wb = sample()
    wb["Finance Input"]["B10"] = False
    wb["Finance Input"]["M25"] = None
    wb.save(path)
    wb.close()
    assert FinancialForecastWorkbookService().analyze(path).existing_finance
