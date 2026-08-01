from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import os
import shutil
import tempfile

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from progress_studio.domain.export_models import ExportResult, ExportValidation
from progress_studio.domain.mapping_models import ActivityRow, AllocationRecord, BOQRow
from progress_studio.infrastructure.excel.xlsx_package_validator import validate_xlsx_tables
from progress_studio.infrastructure.excel.amount_workbook import find_header, normalize_header
from progress_studio.infrastructure.excel.mapping_reader import validate_progress_workbook_contract
from progress_studio.infrastructure.excel.calculation_policy import request_full_excel_recalculation

CURRENCY_FORMAT = '#,##0.00'
PERCENT_FORMAT = '0.00%'
MAPPING_SHEET = 'BOQ Activity Mapping'
SUMMARY_SHEET = 'Mapping Summary'
EXTENSION_SHEET = 'ProgressStudio Extensions'


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
        activities: list[ActivityRow] | None = None,
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
                validate_progress_workbook_contract(workbook)
                totals = self._allocation_totals(boq_rows, allocations)
                activities = activities or []
                supplemental_ids = {row.activity_id for row in activities if row.is_supplemental}
                main_totals = {key: value for key, value in totals.items() if key not in supplemental_ids}
                supplemental_total = sum(value for key, value in totals.items() if key in supplemental_ids)
                amount_rows = self._write_main_amounts(workbook, main_totals, validation.allocated_amount - supplemental_total)
                self._write_amount_mapping(workbook, main_totals)
                self._write_extension_sheet(workbook, activities, totals)
                mapping_rows = self._write_mapping_sheet(workbook, boq_rows, allocations)
                self._write_summary_sheet(workbook, validation, progress_file.name, output_file.name)
                request_full_excel_recalculation(workbook)
                workbook.save(temp_file)
            finally:
                workbook.close()
            validate_xlsx_tables(temp_file)
            _atomic_replace(temp_file, output_file)
            return ExportResult(output_file, validation, amount_rows, mapping_rows)
        except Exception:
            temp_file.unlink(missing_ok=True)
            raise


    @staticmethod
    def _write_extension_sheet(workbook, activities: list[ActivityRow], totals: dict[str, float]) -> None:
        _safe_remove_sheet(workbook, EXTENSION_SHEET)
        supplemental = [row for row in activities if row.is_supplemental]
        if not supplemental:
            return
        ws = workbook.create_sheet(EXTENSION_SHEET)
        ws.append([
            'Origin', 'Parent WBS', 'Supplemental WBS Code', 'Supplemental WBS Name',
            'Activity ID', 'Activity Name', 'Mapped Amount', 'Export Contract',
        ])
        _style_header(ws)
        for row in supplemental:
            ws.append([
                'Progress Studio', row.parent_wbs, row.supplemental_wbs_code,
                row.supplemental_wbs_name, row.activity_id, row.description,
                totals.get(row.activity_id, 0.0),
                'Extension only - not inserted into source schedule hierarchy',
            ])
            ws.cell(ws.max_row, 7).number_format = CURRENCY_FORMAT
        ws.freeze_panes = 'A2'
        for index, width in enumerate((18, 18, 24, 32, 18, 45, 18, 55), start=1):
            ws.column_dimensions[ws.cell(1, index).column_letter].width = width

    @staticmethod
    def _allocation_totals(boq_rows, allocations) -> dict[str, float]:
        by_key = {row.key: row for row in boq_rows}
        totals: dict[str, float] = defaultdict(float)
        for allocation in allocations:
            row = by_key.get(allocation.boq_key)
            if row is None:
                raise ValueError(f'Allocation references missing BOQ item: {allocation.boq_key}')
            totals[allocation.activity_id.strip().upper()] += (
                row.amount * allocation.share_percent / 100.0
            )
        return dict(totals)

    @staticmethod
    def _write_main_amounts(workbook, totals: dict[str, float], expected_total: float) -> int:
        ws = workbook['main']
        header_row, headers = find_header(
            ws, ['Row Type', 'Activity ID', 'P/A', 'Outline Level', 'Amount']
        )
        row_type_col = headers['row type']
        activity_col = headers['activity id']
        pa_col = headers['p/a']
        outline_col = headers['outline level']
        amount_col = headers['amount']

        plan_rows: list[dict[str, object]] = []
        activity_rows: dict[str, int] = {}
        for row_index in range(header_row + 1, ws.max_row + 1):
            if str(ws.cell(row_index, pa_col).value or '').strip().upper() != 'P':
                continue
            row_type = normalize_header(ws.cell(row_index, row_type_col).value)
            if row_type not in {'project summary', 'wbs', 'activity'}:
                continue
            try:
                level = int(ws.cell(row_index, outline_col).value or 0)
            except (TypeError, ValueError):
                raise ValueError(f'Invalid Outline Level in main worksheet row {row_index}.')
            plan_rows.append({'row': row_index, 'row_type': row_type, 'level': level})
            if row_type == 'activity':
                activity_id = str(ws.cell(row_index, activity_col).value or '').strip().upper()
                if not activity_id:
                    raise ValueError(f'Blank Activity ID in main worksheet row {row_index}.')
                if activity_id in activity_rows:
                    raise ValueError(f'Duplicate Activity ID in main worksheet: {activity_id}')
                activity_rows[activity_id] = row_index

        missing = sorted(set(totals) - set(activity_rows))
        if missing:
            preview = ', '.join(missing[:10])
            suffix = '' if len(missing) <= 10 else f' (+{len(missing) - 10} more)'
            raise ValueError('Mapped Activity IDs were not found in main worksheet: ' + preview + suffix)

        for activity_id, row_index in activity_rows.items():
            ws.cell(row_index, amount_col).value = totals.get(activity_id, 0.0)
            ws.cell(row_index, amount_col).number_format = CURRENCY_FORMAT
            if row_index + 1 <= ws.max_row and str(ws.cell(row_index + 1, pa_col).value or '').strip().upper() == 'A':
                ws.cell(row_index + 1, amount_col).value = None

        amount_letter = get_column_letter(amount_col)
        row_type_letter = get_column_letter(row_type_col)
        pa_letter = get_column_letter(pa_col)
        for index in range(len(plan_rows) - 1, -1, -1):
            item = plan_rows[index]
            row_index = int(item['row'])
            row_type = str(item['row_type'])
            level = int(item['level'])
            if row_type not in {'wbs', 'project summary'}:
                continue
            end_row = row_index
            for descendant in plan_rows[index + 1:]:
                if int(descendant['level']) <= level:
                    break
                end_row = int(descendant['row'])
            ws.cell(row_index, amount_col).value = 0 if end_row <= row_index else (
                f'=SUMIFS(${amount_letter}${row_index + 1}:${amount_letter}${end_row},'
                f'${row_type_letter}${row_index + 1}:${row_type_letter}${end_row},"Activity",'
                f'${pa_letter}${row_index + 1}:${pa_letter}${end_row},"P")'
            )
            ws.cell(row_index, amount_col).number_format = CURRENCY_FORMAT
            if row_index + 1 <= ws.max_row and str(ws.cell(row_index + 1, pa_col).value or '').strip().upper() == 'A':
                ws.cell(row_index + 1, amount_col).value = None

        written_total = sum(float(totals.get(activity_id, 0.0)) for activity_id in activity_rows)
        if abs(written_total - expected_total) > 0.01:
            raise ValueError(
                'Export reconciliation failed: main Activity Amount total '
                f'{written_total:,.2f} does not match allocated amount {expected_total:,.2f}.'
            )
        return len(activity_rows)

    @staticmethod
    def _write_amount_mapping(workbook, totals: dict[str, float]) -> None:
        if 'Amount Mapping' not in workbook.sheetnames:
            raise ValueError("Worksheet 'Amount Mapping' was not found in the Progress workbook.")
        ws = workbook['Amount Mapping']
        headers = _headers(ws)
        missing = [name for name in ('activity id', 'amount') if name not in headers]
        if missing:
            raise ValueError('Missing Amount Mapping columns: ' + ', '.join(missing))
        activity_col = headers['activity id']
        amount_col = headers['amount']
        status_col = headers.get('status')
        for row_index in range(2, ws.max_row + 1):
            activity_id = str(ws.cell(row_index, activity_col).value or '').strip().upper()
            if not activity_id:
                continue
            ws.cell(row_index, amount_col).value = totals.get(activity_id, 0.0)
            ws.cell(row_index, amount_col).number_format = CURRENCY_FORMAT
            if status_col:
                ws.cell(row_index, status_col).value = 'MAPPED' if totals.get(activity_id, 0.0) > 0 else 'UNMAPPED'

    @staticmethod
    def _write_mapping_sheet(workbook, boq_rows, allocations) -> int:
        _safe_remove_sheet(workbook, MAPPING_SHEET)
        ws = workbook.create_sheet(MAPPING_SHEET)
        headers = [
            'Activity ID', 'BOQ Key', 'Source Sheet', 'Source Row',
            'WBS-2', 'WBS-3', 'WBS-4', 'BOQ Description', 'BOQ Amount',
            'Share %', 'Allocated Amount', 'Mapping ID', 'BOQ ID',
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
                allocation.activity_id, row.key, row.source_sheet, row.source_row,
                row.wbs2, row.wbs3, row.wbs4, row.description, row.amount,
                allocation.share_percent / 100.0, allocated, mapping_id, row.stable_id or row.key,
            ])
            current = ws.max_row
            ws.cell(current, 9).number_format = CURRENCY_FORMAT
            ws.cell(current, 10).number_format = PERCENT_FORMAT
            ws.cell(current, 11).number_format = CURRENCY_FORMAT

        ws.freeze_panes = 'A2'
        widths = [16, 32, 22, 12, 22, 28, 28, 60, 18, 12, 18, 16, 20]
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
