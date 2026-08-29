from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import os
import shutil
import tempfile
from typing import Callable

from openpyxl import load_workbook

from progress_studio.domain.earned_value import EarnedValueResult
from progress_studio.infrastructure.excel.earned_value_input_reader import (
    EarnedValueInputWorkbookReader,
    EarnedValueWorkbookInputError,
)
from progress_studio.infrastructure.excel.earned_value_workbook import (
    EARNED_VALUE_SHEET,
    render_earned_value_sheet,
)
from progress_studio.infrastructure.excel.rebuild_workbook_reader import (
    RebuildWorkbookReader,
    RebuildWorkbookReadError,
)
from progress_studio.services.earned_value_deriver import (
    EarnedValueDeriver,
    EarnedValueInputError,
)


class EarnedValueRebuildError(ValueError):
    """Raised when a workbook cannot satisfy the EV workspace contract."""


@dataclass(frozen=True, slots=True)
class EarnedValueRebuildAnalysis:
    workbook: Path
    cutoff_date: datetime
    activity_count: int
    boq_count: int
    allocation_count: int
    project_bac: float
    existing_earned_value_sheet: bool


@dataclass(frozen=True, slots=True)
class EarnedValueRebuildResult:
    source_workbook: Path
    output_workbook: Path
    cutoff_date: datetime
    activity_count: int
    boq_count: int
    allocation_count: int
    project_bac: float
    refreshed_sheet: str


def _as_datetime(value: object) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    return None


class EarnedValueRebuildService:
    """Standalone EV boundary used by the Rebuild workspace.

    EV owns only the ``Earned Value`` sheet. It reads:
    - current ``main`` progress values,
    - embedded BOQ/mapping provenance,
    - the management reporting cutoff.

    Progress/Payment rebuild ownership is deliberately untouched.
    """

    def __init__(
        self,
        *,
        input_reader: EarnedValueInputWorkbookReader | None = None,
        main_reader: RebuildWorkbookReader | None = None,
        deriver: EarnedValueDeriver | None = None,
        renderer: Callable[[object, EarnedValueResult], None] = render_earned_value_sheet,
    ) -> None:
        self.input_reader = input_reader or EarnedValueInputWorkbookReader()
        self.main_reader = main_reader or RebuildWorkbookReader()
        self.deriver = deriver or EarnedValueDeriver()
        self.renderer = renderer

    def analyze(self, workbook_path: Path) -> EarnedValueRebuildAnalysis:
        path = self._validate_path(workbook_path)
        result, activity_count, boq_count, allocation_count = self._derive(path)
        return EarnedValueRebuildAnalysis(
            workbook=path,
            cutoff_date=self._require_cutoff(path),
            activity_count=activity_count,
            boq_count=boq_count,
            allocation_count=allocation_count,
            project_bac=float(result.project_bac),
            existing_earned_value_sheet=self._has_sheet(path, EARNED_VALUE_SHEET),
        )

    def generate(
        self,
        source_workbook: Path,
        output_workbook: Path,
    ) -> EarnedValueRebuildResult:
        source = self._validate_path(source_workbook)
        output = Path(output_workbook).expanduser().resolve()
        if output.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise EarnedValueRebuildError(
                "Earned Value output must use .xlsx or .xlsm."
            )
        if source == output:
            raise EarnedValueRebuildError(
                "Earned Value output must be a new workbook path."
            )

        result, activity_count, boq_count, allocation_count = self._derive(source)
        cutoff = self._require_cutoff(source)

        output.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            prefix=f".{output.stem}.ev4.",
            suffix=output.suffix,
            dir=output.parent,
        )
        os.close(fd)
        temp_path = Path(temp_name)
        keep_vba = source.suffix.lower() == ".xlsm"

        try:
            shutil.copy2(source, temp_path)
            workbook = load_workbook(
                temp_path,
                read_only=False,
                data_only=False,
                keep_vba=keep_vba,
            )
            try:
                self.renderer(workbook, result)
                workbook.save(temp_path)
            finally:
                workbook.close()
            os.replace(temp_path, output)
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise

        return EarnedValueRebuildResult(
            source_workbook=source,
            output_workbook=output,
            cutoff_date=cutoff,
            activity_count=activity_count,
            boq_count=boq_count,
            allocation_count=allocation_count,
            project_bac=float(result.project_bac),
            refreshed_sheet=EARNED_VALUE_SHEET,
        )

    def _derive(
        self,
        path: Path,
    ) -> tuple[EarnedValueResult, int, int, int]:
        try:
            embedded = self.input_reader.read(path)
            dataset = self.main_reader.read_main_dataset(path)
            cutoff = self._require_cutoff(path)
            result = self.deriver.derive(
                dataset,
                embedded.boq_rows,
                embedded.allocations,
                cutoff_date=cutoff,
            )
        except (
            EarnedValueWorkbookInputError,
            RebuildWorkbookReadError,
            EarnedValueInputError,
            KeyError,
            ValueError,
        ) as exc:
            if isinstance(exc, EarnedValueRebuildError):
                raise
            raise EarnedValueRebuildError(str(exc)) from exc

        activity_count = len(tuple(dataset.activities))
        if activity_count <= 0:
            raise EarnedValueRebuildError(
                "Worksheet 'main' contains no valid Plan Activity rows."
            )
        if not result.project_points:
            raise EarnedValueRebuildError(
                "Earned Value has no reporting periods to render."
            )
        return (
            result,
            activity_count,
            len(embedded.boq_rows),
            len(embedded.allocations),
        )

    def _require_cutoff(self, path: Path) -> datetime:
        cutoff = self._read_cutoff(path)
        if cutoff is None:
            raise EarnedValueRebuildError(
                "Earned Value requires a reporting cutoff date. "
                "Save a valid cutoff in Dashboard!K5 or the main Cutoff Date control."
            )
        return cutoff

    @staticmethod
    def _read_cutoff(path: Path) -> datetime | None:
        keep_vba = path.suffix.lower() == ".xlsm"
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True,
            keep_vba=keep_vba,
        )
        try:
            # EV is a management view, so the Dashboard reporting cutoff is the
            # authoritative first choice. The main local control is a fallback
            # for older/partially rebuilt workbooks.
            if "Dashboard" in workbook.sheetnames:
                cutoff = _as_datetime(workbook["Dashboard"]["K5"].value)
                if cutoff is not None:
                    return cutoff

            if "main" in workbook.sheetnames:
                ws = workbook["main"]
                max_row = min(ws.max_row, 80)
                max_col = min(ws.max_column, 30)
                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        value = str(ws.cell(row, col).value or "").strip().lower()
                        if value != "cutoff date":
                            continue
                        if col < max_col:
                            cutoff = _as_datetime(ws.cell(row, col + 1).value)
                            if cutoff is not None:
                                return cutoff
            return None
        finally:
            workbook.close()

    @staticmethod
    def _has_sheet(path: Path, sheet_name: str) -> bool:
        keep_vba = path.suffix.lower() == ".xlsm"
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=False,
            keep_vba=keep_vba,
        )
        try:
            return sheet_name in workbook.sheetnames
        finally:
            workbook.close()

    @staticmethod
    def _validate_path(workbook_path: Path) -> Path:
        path = Path(workbook_path).expanduser().resolve()
        if not path.is_file():
            raise EarnedValueRebuildError(
                f"Earned Value workbook was not found: {path}"
            )
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise EarnedValueRebuildError(
                "Select a Progress Studio .xlsx or .xlsm workbook."
            )
        return path
