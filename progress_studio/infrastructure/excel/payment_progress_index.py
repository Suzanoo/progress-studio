from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from progress_studio.domain.payment_models import (
    ActivityProgress,
    ActivityProgressBucket,
    ActivityProgressIndex,
)
from progress_studio.infrastructure.excel.payment_workbook import (
    MAIN_NS,
    PaymentWorkbookError,
    PaymentWorkbookSnapshotter,
)


class ActivityProgressIndexReader:
    """Build one compact cumulative-plan index from the current ``main`` sheet."""

    MAIN_SHEET = "main"
    HEADER_SCAN_LIMIT = 30
    EPSILON = 1e-12

    def read(self, workbook_path: Path) -> ActivityProgressIndex:
        path = Path(workbook_path)
        if not path.exists():
            raise PaymentWorkbookError(f"Progress workbook was not found: {path}")
        try:
            with ZipFile(path, "r") as package:
                sheet_map = PaymentWorkbookSnapshotter._sheet_map(package)
                if self.MAIN_SHEET not in sheet_map:
                    raise PaymentWorkbookError("Worksheet 'main' was not found.")
                root = ET.fromstring(package.read(sheet_map[self.MAIN_SHEET]))
                shared_strings = PaymentWorkbookSnapshotter._shared_strings(package)
                return self._parse(path, root, shared_strings)
        except PaymentWorkbookError:
            raise
        except Exception as exc:
            raise PaymentWorkbookError(f"Activity progress index could not be built: {exc}") from exc

    def _parse(self, path: Path, root: ET.Element, shared_strings: list[str]) -> ActivityProgressIndex:
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise PaymentWorkbookError("Worksheet 'main' has no data.")

        rows = sheet_data.findall(f"{{{MAIN_NS}}}row")
        header_row, header_columns = self._find_headers(rows, shared_strings)
        timescale = self._timescale_columns(rows, header_row, header_columns, shared_strings)
        if not timescale:
            raise PaymentWorkbookError("Weekly timescale columns were not found in 'main'.")

        activities: dict[str, ActivityProgress] = {}
        duplicate_ids: list[str] = []
        for row in rows:
            row_num = int(row.attrib.get("r", "0"))
            if row_num <= header_row:
                continue
            cells = self._row_cells(row)
            row_type = self._cell_text(cells.get(header_columns["row type"]), shared_strings).strip().lower()
            pa = self._cell_text(cells.get(header_columns["p/a"]), shared_strings).strip().upper()
            if row_type != "activity" or pa != "P":
                continue
            activity_id = self._cell_text(cells.get(header_columns["activity id"]), shared_strings).strip()
            if not activity_id:
                continue
            if activity_id in activities:
                duplicate_ids.append(activity_id)
                continue

            raw: list[tuple[int, str, date, float]] = []
            total = 0.0
            for col_idx, col_letters, week_start in timescale:
                value = self._numeric(cells.get(col_idx), shared_strings)
                if value is None or abs(value) <= self.EPSILON:
                    continue
                if value < -self.EPSILON:
                    raise PaymentWorkbookError(
                        f"Negative plan distribution found for {activity_id} at {col_letters}{row_num}."
                    )
                total += value
                raw.append((col_idx, col_letters, week_start, value))

            buckets: list[ActivityProgressBucket] = []
            if total > self.EPSILON:
                cumulative = 0.0
                for col_idx, col_letters, week_start, value in raw:
                    normalized = value / total
                    cumulative = min(cumulative + normalized, 1.0)
                    buckets.append(
                        ActivityProgressBucket(
                            column_index=col_idx,
                            column_letter=col_letters,
                            week_start=week_start,
                            incremental_fraction=normalized,
                            cumulative_fraction=cumulative,
                        )
                    )
                if buckets:
                    last = buckets[-1]
                    buckets[-1] = ActivityProgressBucket(
                        column_index=last.column_index,
                        column_letter=last.column_letter,
                        week_start=last.week_start,
                        incremental_fraction=last.incremental_fraction,
                        cumulative_fraction=1.0,
                    )

            activities[activity_id] = ActivityProgress(
                activity_id=activity_id,
                row_number=row_num,
                buckets=tuple(buckets),
            )

        if duplicate_ids:
            duplicates = ", ".join(sorted(set(duplicate_ids))[:10])
            raise PaymentWorkbookError(f"Duplicate Plan Activity IDs in main: {duplicates}")
        if not activities:
            raise PaymentWorkbookError("No Plan activity rows were found in 'main'.")

        return ActivityProgressIndex(
            workbook=path,
            sheet=self.MAIN_SHEET,
            timescale_columns=tuple(timescale),
            activities=activities,
        )

    def _find_headers(self, rows: list[ET.Element], shared_strings: list[str]) -> tuple[int, dict[str, int]]:
        required = {"row type", "p/a", "activity id"}
        for row in rows:
            row_num = int(row.attrib.get("r", "0"))
            if row_num > self.HEADER_SCAN_LIMIT:
                break
            columns: dict[str, int] = {}
            for col_idx, cell in self._row_cells(row).items():
                text = self._cell_text(cell, shared_strings).strip().lower()
                if text in required or text == "xml amount":
                    columns[text] = col_idx
            if required.issubset(columns):
                return row_num, columns
        raise PaymentWorkbookError("Required main headers Row Type / P/A / Activity ID were not found.")

    def _timescale_columns(
        self,
        rows: list[ET.Element],
        header_row: int,
        header_columns: dict[str, int],
        shared_strings: list[str],
    ) -> list[tuple[int, str, date]]:
        row = next((r for r in rows if int(r.attrib.get("r", "0")) == header_row), None)
        if row is None:
            return []
        fixed_end = header_columns.get("xml amount", max(header_columns.values()))
        result: list[tuple[int, str, date]] = []
        for col_idx, cell in sorted(self._row_cells(row).items()):
            if col_idx <= fixed_end:
                continue
            week_start = self._excel_date(cell, shared_strings)
            if week_start is not None:
                result.append((col_idx, self._column_letters(col_idx), week_start))
        return result

    @staticmethod
    def _row_cells(row: ET.Element) -> dict[int, ET.Element]:
        result: dict[int, ET.Element] = {}
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            match = re.match(r"([A-Z]+)", cell.attrib.get("r", ""))
            if match:
                result[ActivityProgressIndexReader._column_index(match.group(1))] = cell
        return result

    @staticmethod
    def _column_index(letters: str) -> int:
        value = 0
        for ch in letters:
            value = value * 26 + ord(ch) - 64
        return value

    @staticmethod
    def _column_letters(index: int) -> str:
        chars: list[str] = []
        while index:
            index, rem = divmod(index - 1, 26)
            chars.append(chr(65 + rem))
        return "".join(reversed(chars))

    @staticmethod
    def _cell_text(cell: ET.Element | None, shared_strings: list[str]) -> str:
        if cell is None:
            return ""
        return PaymentWorkbookSnapshotter._cell_text(cell, shared_strings)

    @classmethod
    def _numeric(cls, cell: ET.Element | None, shared_strings: list[str]) -> float | None:
        if cell is None:
            return None
        text = cls._cell_text(cell, shared_strings).strip()
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None

    @classmethod
    def _excel_date(cls, cell: ET.Element | None, shared_strings: list[str]) -> date | None:
        text = cls._cell_text(cell, shared_strings).strip() if cell is not None else ""
        if not text:
            return None
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(text))).date()
        except ValueError:
            try:
                return datetime.fromisoformat(text).date()
            except ValueError:
                return None
