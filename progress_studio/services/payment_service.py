from __future__ import annotations

from pathlib import Path

from progress_studio.domain.payment_models import PaymentSnapshotResult, PaymentWorkbookValidation
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookSnapshotter


class PaymentService:
    def __init__(self, snapshotter: PaymentWorkbookSnapshotter | None = None) -> None:
        self.snapshotter = snapshotter or PaymentWorkbookSnapshotter()

    def validate_workbook(self, workbook: Path) -> PaymentWorkbookValidation:
        return self.snapshotter.validate(Path(workbook))

    def create_payment_snapshot(self, source: Path, output: Path) -> PaymentSnapshotResult:
        return self.snapshotter.create_snapshot(Path(source), Path(output))
