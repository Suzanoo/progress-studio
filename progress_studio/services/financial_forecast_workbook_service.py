"""FF-2 workbook-facing entry point; no schedule inference or FF-3 UI."""
import os
from pathlib import Path
import tempfile
from openpyxl import load_workbook
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900
from progress_studio.infrastructure.excel.finance_input_workbook import create_inputs, read_finance_inputs
from progress_studio.infrastructure.excel.financial_forecast_workbook import build_finance_view
from progress_studio.infrastructure.excel.finance_package import merge_finance_parts
from progress_studio.infrastructure.excel.xlsx_package_validator import validate_xlsx_tables


class FinancialForecastWorkbookService:
    def generate(self, source_workbook, output_workbook, *, opening_date=None,
                 actuals_through=None, currency=None, capacity_rows=None):
        source=Path(source_workbook).expanduser().resolve()
        output=Path(output_workbook).expanduser().resolve()
        if source.suffix.lower()!='.xlsx' or output.suffix.lower()!='.xlsx':
            raise ValueError('Finance V1 requires .xlsx workbooks')
        if source==output or output.exists():raise ValueError('Finance output must be a new workbook path')
        if not source.is_file():raise FileNotFoundError(source)
        output.parent.mkdir(parents=True,exist_ok=True)
        fd,name=tempfile.mkstemp(prefix='.finance.',suffix='.xlsx',dir=output.parent)
        os.close(fd);temp=Path(name)
        try:
            wb=load_workbook(source,keep_links=True)
            try:
                if wb.epoch != CALENDAR_WINDOWS_1900:
                    raise ValueError("Finance V1 requires the Excel 1900 date system")
                create_inputs(wb,opening_date=opening_date,actuals_through=actuals_through,
                              currency=currency,capacity_rows=capacity_rows if capacity_rows is not None else 1 if "Finance Input" in wb.sheetnames else 100)
                read_finance_inputs(wb)  # FF-1 reference validation; never save invalid edits.
                build_finance_view(wb)
                wb.save(temp)
            finally:wb.close()
            merge_finance_parts(source,temp)
            validate_xlsx_tables(temp)
            # Reopen the published package, not only the intermediate serializer.
            check=load_workbook(temp,keep_links=True)
            try:read_finance_inputs(check)
            finally:check.close()
            os.replace(temp,output)
        finally:temp.unlink(missing_ok=True)
        return output
