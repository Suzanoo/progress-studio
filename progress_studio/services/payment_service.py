from __future__ import annotations

from pathlib import Path

from progress_studio.domain.payment_models import (
    PaymentInputResult,
    PaymentInputValidation,
    PaymentSnapshotResult,
    PaymentWorkbookValidation,
)
from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookSnapshotter


class PaymentService:
    def __init__(
        self,
        snapshotter: PaymentWorkbookSnapshotter | None = None,
        payment_input: PaymentInputWorkbook | None = None,
    ) -> None:
        self.snapshotter = snapshotter or PaymentWorkbookSnapshotter()
        self.payment_input = payment_input or PaymentInputWorkbook()

    def validate_workbook(self, workbook: Path) -> PaymentWorkbookValidation:
        return self.snapshotter.validate(Path(workbook))

    def create_payment_snapshot(self, source: Path, output: Path) -> PaymentSnapshotResult:
        return self.snapshotter.create_snapshot(Path(source), Path(output))

    def create_fake_payment_input(self, source: Path, output: Path, periods: int) -> PaymentInputResult:
        return self.payment_input.create(Path(source), Path(output), periods)

    def validate_payment_input(self, payment_workbook: Path, progress_workbook: Path | None = None) -> PaymentInputValidation:
        progress = Path(progress_workbook) if progress_workbook is not None else None
        return self.payment_input.validate(Path(payment_workbook), progress)
