from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from progress_studio.domain.payment_models import (
    PaymentInputData,
    PaymentPeriodRequirements,
    PaymentRequirement,
)
from progress_studio.infrastructure.excel.payment_workbook import (
    MAIN_NS,
    PaymentWorkbookError,
    PaymentWorkbookSnapshotter,
)


class PaymentInputSparseReader:
    """Read the user Payment Input workbook once and keep only populated requirements.

    Blank payment cells are intentionally absent from the returned model. Explicit
    0% cells remain requirements, so ``blank != 0%`` is preserved without carrying
    a dense Activity x Payment matrix in memory.
    """

    SHEET = "Payment Input"
    HEADER_ROW = 6
    DATE_ROW = 7
    FIRST_ACTIVITY_ROW = 8

    def read(self, workbook_path: Path) -> PaymentInputData:
        path = Path(workbook_path)
        if not path.exists():
            raise PaymentWorkbookError(f"Payment workbook was not found: {path}")
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise PaymentWorkbookError("Select a Payment Requirement .xlsx or .xlsm workbook.")

        try:
            with ZipFile(path, "r") as package:
                sheet_map = PaymentWorkbookSnapshotter._sheet_map(package)
                if self.SHEET not in sheet_map:
                    raise PaymentWorkbookError("Worksheet 'Payment Input' was not found.")
                root = ET.fromstring(package.read(sheet_map[self.SHEET]))
                shared_strings = PaymentWorkbookSnapshotter._shared_strings(package)
                return self._parse(path, root, shared_strings)
        except PaymentWorkbookError:
            raise
        except Exception as exc:
            raise PaymentWorkbookError(f"Payment Input cannot be opened: {exc}") from exc

    def _parse(self, path: Path, root: ET.Element, shared_strings: list[str]) -> PaymentInputData:
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise PaymentWorkbookError("Payment Input worksheet has no data.")

        rows = {
            int(row.attrib.get("r", "0")): row
            for row in sheet_data.findall(f"{{{MAIN_NS}}}row")
        }
        header = rows.get(self.HEADER_ROW)
        if header is None:
            raise PaymentWorkbookError("Payment Input header row was not found.")

        header_cells = self._row_cells(header)
        if self._cell_text(header_cells.get(1), shared_strings).strip().lower() != "activity id":
            raise PaymentWorkbookError("Payment Input header 'Activity ID' was not found.")

        periods: list[tuple[int, str]] = []
        seen_periods: set[str] = set()
        for col, cell in sorted(header_cells.items()):
            text = self._cell_text(cell, shared_strings).strip().upper()
            if not re.fullmatch(r"P\d+", text):
                continue
            if text in seen_periods:
                raise PaymentWorkbookError(f"Duplicate payment period header: {text}")
            seen_periods.add(text)
            periods.append((col, text))
        if not periods:
            raise PaymentWorkbookError("No payment period columns were found.")

        date_cells = self._row_cells(rows.get(self.DATE_ROW)) if rows.get(self.DATE_ROW) is not None else {}
        period_dates = {
            period_id: self._cell_date(date_cells.get(col_idx), shared_strings)
            for col_idx, period_id in periods
        }

        requirements: dict[str, list[PaymentRequirement]] = {period_id: [] for _, period_id in periods}
        activity_ids: list[str] = []
        seen_activities: set[str] = set()

        for row_num in sorted(r for r in rows if r >= self.FIRST_ACTIVITY_ROW):
            row = rows[row_num]
            cells = self._row_cells(row)
            activity_id = self._cell_text(cells.get(1), shared_strings).strip()
            if not activity_id:
                continue
            if activity_id in seen_activities:
                raise PaymentWorkbookError(f"Duplicate Activity ID in Payment Input: {activity_id}")
            seen_activities.add(activity_id)
            activity_ids.append(activity_id)

            for col_idx, period_id in periods:
                cell = cells.get(col_idx)
                if cell is None or self._is_blank(cell, shared_strings):
                    continue
                value = self._requirement_value(cell, shared_strings)
                if value is None:
                    continue
                if value < 0.0 or value > 1.0:
                    ref = cell.attrib.get("r", f"row {row_num}")
                    raise PaymentWorkbookError(
                        f"Payment requirement at {ref} must be between 0% and 100%."
                    )
                requirements[period_id].append(
                    PaymentRequirement(
                        activity_id=activity_id,
                        required_fraction=value,
                        source_row=row_num,
                        source_column=col_idx,
                    )
                )

        if not activity_ids:
            raise PaymentWorkbookError("No Activity IDs were found in Payment Input.")

        period_models = tuple(
            PaymentPeriodRequirements(
                period_id=period_id,
                column_index=col_idx,
                payment_date=period_dates[period_id],
                requirements=tuple(requirements[period_id]),
            )
            for col_idx, period_id in periods
        )
        return PaymentInputData(
            workbook=path,
            sheet=self.SHEET,
            periods=period_models,
            activity_ids=tuple(activity_ids),
            populated_requirements=sum(len(period.requirements) for period in period_models),
        )

    @staticmethod
    def _row_cells(row: ET.Element | None) -> dict[int, ET.Element]:
        if row is None:
            return {}
        result: dict[int, ET.Element] = {}
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            match = re.match(r"([A-Z]+)", cell.attrib.get("r", ""))
            if match:
                result[PaymentInputSparseReader._column_index(match.group(1))] = cell
        return result

    @staticmethod
    def _column_index(letters: str) -> int:
        value = 0
        for ch in letters:
            value = value * 26 + ord(ch) - 64
        return value

    @staticmethod
    def _cell_text(cell: ET.Element | None, shared_strings: list[str]) -> str:
        if cell is None:
            return ""
        return PaymentWorkbookSnapshotter._cell_text(cell, shared_strings)

    @classmethod
    def _is_blank(cls, cell: ET.Element, shared_strings: list[str]) -> bool:
        return cls._cell_text(cell, shared_strings).strip() == ""

    @classmethod
    def _requirement_value(cls, cell: ET.Element, shared_strings: list[str]) -> float | None:
        text = cls._cell_text(cell, shared_strings).strip()
        if not text:
            return None
        if text.endswith("%"):
            try:
                return float(text[:-1].strip()) / 100.0
            except ValueError as exc:
                raise PaymentWorkbookError(f"Invalid percentage value at {cell.attrib.get('r', '?')}: {text}") from exc
        try:
            return float(text)
        except ValueError as exc:
            raise PaymentWorkbookError(
                f"Payment requirement at {cell.attrib.get('r', '?')} must be a percentage value."
            ) from exc

    @classmethod
    def _cell_date(cls, cell: ET.Element | None, shared_strings: list[str]) -> date | None:
        if cell is None:
            return None
        text = cls._cell_text(cell, shared_strings).strip()
        if not text:
            return None
        try:
            serial = float(text)
            return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
        except ValueError:
            try:
                return datetime.fromisoformat(text).date()
            except ValueError:
                return None
