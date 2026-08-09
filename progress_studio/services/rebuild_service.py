from __future__ import annotations

from pathlib import Path

from progress_studio.infrastructure.excel.rebuild_workbook_reader import (
    RebuildWorkbookReadError,
    RebuildWorkbookReader,
)

from progress_studio.domain.rebuild_models import (
    RebuildMode,
    RebuildSheetContract,
    RebuildWorkbookAnalysis,
)


class RebuildContractError(ValueError):
    """Raised when an input workbook cannot satisfy the standalone rebuild contract."""


DEFAULT_REBUILD_CONTRACT = RebuildSheetContract(
    source_of_truth=("main", "Payment Input"),
    preserve=("main", "Payment Input"),
    generated_progress=(
        "main_monthly",
        "progress",
        "progress_table",
        "Dashboard_Data",
        "Dashboard",
    ),
    generated_payment=("Payment",),
    internal_preserve=(
        "Info",
        "Timescale Info",
        "Amount Mapping",
        "Distribution Report",
        "ProgressStudio Extensions",
        "BOQ Activity Mapping",
        "Mapping Summary",
        "Rebuild Audit",
    ),
)


class WorkbookRebuildEngine:
    """Standalone rebuild boundary.

    This engine deliberately knows nothing about XML, BOQ files, .progressstudio
    sessions, .boqstudio sessions, mapping allocations, or WorkingScheduleTree.

    `main` in the supplied workbook is authoritative. Payment mode additionally
    requires the embedded `Payment Input` sheet.

    MS-RB1 only analyzes and plans ownership. It does not regenerate sheets yet.
    """

    MAIN_SHEET = "main"
    PAYMENT_INPUT_SHEET = "Payment Input"

    def __init__(
        self,
        contract: RebuildSheetContract | None = None,
        reader: RebuildWorkbookReader | None = None,
    ) -> None:
        self.contract = contract or DEFAULT_REBUILD_CONTRACT
        self.reader = reader or RebuildWorkbookReader()

    def analyze(
        self,
        workbook_path: Path,
        mode: RebuildMode | str,
    ) -> RebuildWorkbookAnalysis:
        path = Path(workbook_path).expanduser().resolve()
        rebuild_mode = self._mode(mode)
        self._validate_path(path)

        try:
            probe = self.reader.probe(path)
        except RebuildWorkbookReadError as exc:
            raise RebuildContractError(str(exc)) from exc

        if probe.activity_count <= 0:
            raise RebuildContractError(
                "Worksheet 'main' contains no valid Plan Activity rows."
            )

        payment_input_present = self.PAYMENT_INPUT_SHEET in probe.sheet_names
        if rebuild_mode is RebuildMode.PAYMENT and not payment_input_present:
            raise RebuildContractError(
                "Rebuild Payment requires embedded worksheet 'Payment Input'."
            )

        generated = self.contract.generated_for(rebuild_mode)
        existing_generated = tuple(name for name in generated if name in probe.sheet_names)
        missing_generated = tuple(name for name in generated if name not in probe.sheet_names)

        preserve_names = set(self.contract.preserve) | set(self.contract.internal_preserve)
        preserve_present = tuple(
            name for name in probe.sheet_names if name in preserve_names
        )

        known = (
            set(self.contract.preserve)
            | set(self.contract.internal_preserve)
            | set(self.contract.generated_progress)
            | set(self.contract.generated_payment)
        )
        unknown = tuple(name for name in probe.sheet_names if name not in known)

        return RebuildWorkbookAnalysis(
            workbook=path,
            mode=rebuild_mode,
            main_sheet=self.MAIN_SHEET,
            main_rows=probe.main_rows,
            main_columns=probe.main_columns,
            activity_count=probe.activity_count,
            payment_input_present=payment_input_present,
            existing_generated_sheets=existing_generated,
            missing_generated_sheets=missing_generated,
            preserve_sheets_present=preserve_present,
            unknown_sheets=unknown,
            contract=self.contract,
        )

    def generated_sheets_for(self, mode: RebuildMode | str) -> tuple[str, ...]:
        """Return the only sheets a future rebuild execution may replace."""
        return self.contract.generated_for(self._mode(mode))

    @staticmethod
    def _mode(mode: RebuildMode | str) -> RebuildMode:
        if isinstance(mode, RebuildMode):
            return mode
        try:
            return RebuildMode(str(mode).strip().lower())
        except ValueError as exc:
            raise RebuildContractError(
                "Rebuild mode must be 'progress' or 'payment'."
            ) from exc

    @staticmethod
    def _validate_path(path: Path) -> None:
        if not path.is_file():
            raise RebuildContractError(f"Workbook was not found: {path}")
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise RebuildContractError(
                "Standalone rebuild accepts an Excel .xlsx or .xlsm workbook."
            )

