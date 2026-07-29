from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import os
import shutil
import tempfile

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from progress_studio.domain.export_models import ExportResult, ExportValidation
from progress_studio.domain.mapping_models import AllocationRecord, BOQRow

CURRENCY_FORMAT = '#,##0.00'
PERCENT_FORMAT = '0.00%'
MAPPING_SHEET = 'BOQ Activity Mapping'
SUMMARY_SHEET = 'Mapping Summary'


def _headers(ws, row: int = 1) -> dict[str, int]:
    return {
        str(ws.cell(row, col).value or '').strip().lower(): col
        for col in range(1, ws.max_column + 1)
        if str(ws.cell(row, col).value or '').strip()
    }


def _safe_remove_sheet(workbook, name: str) -> None:
    if name in workbook.sheetnames:
        del workbook[name]


def _style_header(ws, row: int = 1) -> None:
    for cell in ws[row]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='4472C4')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def _atomic_replace(temp_file: Path, output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    os.replace(temp_file, output_file)


class MappedWorkbookExporter:
    """Create a final mapped workbook without reading data from GUI widgets."""

    def export(
        self,
        progress_file: Path,
        output_file: Path,
        boq_rows: list[BOQRow],
        allocations: list[AllocationRecord],
        validation: ExportValidation,
        *,
        overwrite: bool = False,
    ) -> ExportResult:
        progress_file = Path(progress_file).resolve()
        output_file = Path(output_file).resolve()
        if not progress_file.is_file():
            raise ValueError(f'Progress workbook was not found: {progress_file}')
        if progress_file == output_file:
            raise ValueError('Export output must be different from the loaded Progress workbook.')
        if output_file.exists() and not overwrite:
            raise FileExistsError(f'Export file already exists: {output_file}')
        if output_file.suffix.lower() != '.xlsx':
            raise ValueError('Export filename must use the .xlsx extension.')

        output_file.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            prefix=f'.{output_file.stem}.', suffix='.tmp.xlsx', dir=output_file.parent
        )
        os.close(fd)
        temp_file = Path(temp_name)
        try:
            shutil.copy2(progress_file, temp_file)
            workbook = load_workbook(temp_file)
            try:
                amount_rows = self._write_amount_mapping(workbook, boq_rows, allocations)
                mapping_rows = self._write_mapping_sheet(workbook, boq_rows, allocations)
                self._write_summary_sheet(workbook, validation, progress_file.name, output_file.name)
                workbook.calculation.calcMode = 'auto'
                workbook.calculation.fullCalcOnLoad = True
                workbook.calculation.forceFullCalc = True
                workbook.save(temp_file)
            finally:
                workbook.close()
            _atomic_replace(temp_file, output_file)
            return ExportResult(output_file, validation, amount_rows, mapping_rows)
        except Exception:
            temp_file.unlink(missing_ok=True)
            raise

    @staticmethod
    def _write_amount_mapping(workbook, boq_rows, allocations) -> int:
        if 'Amount Mapping' not in workbook.sheetnames:
            raise ValueError("Worksheet 'Amount Mapping' was not found in the Progress workbook.")
        ws = workbook['Amount Mapping']
        headers = _headers(ws)
        missing = [name for name in ('activity id', 'amount') if name not in headers]
        if missing:
            raise ValueError('Missing Amount Mapping columns: ' + ', '.join(missing))

        by_key = {row.key: row for row in boq_rows}
        totals: dict[str, float] = defaultdict(float)
        for allocation in allocations:
            row = by_key.get(allocation.boq_key)
            if row is None:
                raise ValueError(f'Allocation references missing BOQ item: {allocation.boq_key}')
            totals[allocation.activity_id] += row.amount * allocation.share_percent / 100.0

        updated = 0
        activity_col = headers['activity id']
        amount_col = headers['amount']
        status_col = headers.get('status')
        for row_index in range(2, ws.max_row + 1):
            activity_id = str(ws.cell(row_index, activity_col).value or '').strip()
            if not activity_id:
                continue
            ws.cell(row_index, amount_col).value = totals.get(activity_id, 0.0)
            ws.cell(row_index, amount_col).number_format = CURRENCY_FORMAT
            if status_col:
                ws.cell(row_index, status_col).value = 'MAPPED' if activity_id in totals else 'UNMAPPED'
            updated += 1
        return updated

    @staticmethod
    def _write_mapping_sheet(workbook, boq_rows, allocations) -> int:
        _safe_remove_sheet(workbook, MAPPING_SHEET)
        ws = workbook.create_sheet(MAPPING_SHEET)
        headers = [
            'Mapping ID', 'Activity ID', 'BOQ ID', 'BOQ Key', 'Source Sheet', 'Source Row',
            'WBS-2', 'WBS-3', 'WBS-4', 'BOQ Description', 'BOQ Amount',
            'Share %', 'Allocated Amount',
        ]
        ws.append(headers)
        _style_header(ws)
        by_key = {row.key: row for row in boq_rows}
        for index, allocation in enumerate(sorted(allocations, key=lambda x: (x.activity_id, x.boq_key)), start=1):
            row = by_key.get(allocation.boq_key)
            if row is None:
                raise ValueError(f'Allocation references missing BOQ item: {allocation.boq_key}')
            allocated = row.amount * allocation.share_percent / 100.0
            mapping_id = f'MAP-{index:06d}'
            ws.append([
                mapping_id, allocation.activity_id, row.stable_id, row.key, row.source_sheet,
                row.source_row, row.wbs2, row.wbs3, row.wbs4, row.description,
                row.amount, allocation.share_percent / 100.0, allocated,
            ])
            current = ws.max_row
            ws.cell(current, 11).number_format = CURRENCY_FORMAT
            ws.cell(current, 12).number_format = PERCENT_FORMAT
            ws.cell(current, 13).number_format = CURRENCY_FORMAT

        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f'A1:M{max(1, ws.max_row)}'
        widths = [16, 16, 20, 32, 22, 12, 22, 28, 28, 60, 18, 12, 18]
        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(1, index).column_letter].width = width
        if ws.max_row >= 2:
            table = Table(displayName='BOQActivityMappingTable', ref=f'A1:M{ws.max_row}')
            table.tableStyleInfo = TableStyleInfo(
                name='TableStyleMedium2', showFirstColumn=False, showLastColumn=False,
                showRowStripes=True, showColumnStripes=False,
            )
            ws.add_table(table)
        return len(allocations)

    @staticmethod
    def _write_summary_sheet(workbook, validation, source_name: str, output_name: str) -> None:
        _safe_remove_sheet(workbook, SUMMARY_SHEET)
        ws = workbook.create_sheet(SUMMARY_SHEET, 0)
        ws['A1'] = 'Progress Studio Mapping Reconciliation'
        ws['A1'].font = Font(size=14, bold=True, color='FFFFFF')
        ws['A1'].fill = PatternFill('solid', fgColor='1F4E78')
        ws.merge_cells('A1:B1')
        rows = [
            ('Source workbook', source_name),
            ('Export workbook', output_name),
            ('Export status', 'Complete' if validation.is_complete else 'Partial'),
            ('Activities', validation.activity_count),
            ('BOQ items', validation.boq_count),
            ('Allocation records', validation.allocation_count),
            ('Mapped activities', validation.mapped_activity_count),
            ('Fully allocated BOQ items', validation.full_boq_count),
            ('Partially allocated BOQ items', validation.partial_boq_count),
            ('Unmapped BOQ items', validation.unmapped_boq_count),
            ('Total BOQ amount', validation.total_boq_amount),
            ('Allocated amount', validation.allocated_amount),
            ('Remaining amount', validation.remaining_amount),
            ('Allocated percent', validation.allocated_percent / 100.0),
        ]
        for row_index, (label, value) in enumerate(rows, start=3):
            ws.cell(row_index, 1).value = label
            ws.cell(row_index, 1).font = Font(bold=True)
            ws.cell(row_index, 2).value = value
        for row_index in (13, 14, 15):
            ws.cell(row_index, 2).number_format = CURRENCY_FORMAT
        ws.cell(16, 2).number_format = PERCENT_FORMAT
        ws.column_dimensions['A'].width = 34
        ws.column_dimensions['B'].width = 28
        ws.freeze_panes = 'A3'
