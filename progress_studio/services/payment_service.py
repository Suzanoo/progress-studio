from __future__ import annotations

from pathlib import Path

from progress_studio.domain.payment_models import (
    PaymentInputData,
    PaymentInputResult,
    PaymentInputValidation,
    PaymentLineRenderResult,
    PaymentMultiLineRenderResult,
    PaymentPositionResult,
    PaymentPreparationResult,
    PaymentSnapshotResult,
    PaymentWorkbookValidation,
)
from progress_studio.infrastructure.excel.payment_input_reader import PaymentInputSparseReader
from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook
from progress_studio.infrastructure.excel.payment_line_renderer import PaymentLineRenderer
from progress_studio.infrastructure.excel.payment_progress_index import ActivityProgressIndexReader
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookSnapshotter
from progress_studio.services.payment_position_engine import PaymentPositionEngine


class PaymentService:
    def __init__(
        self,
        snapshotter: PaymentWorkbookSnapshotter | None = None,
        payment_input: PaymentInputWorkbook | None = None,
        payment_reader: PaymentInputSparseReader | None = None,
        progress_index_reader: ActivityProgressIndexReader | None = None,
        position_engine: PaymentPositionEngine | None = None,
        line_renderer: PaymentLineRenderer | None = None,
    ) -> None:
        self.snapshotter = snapshotter or PaymentWorkbookSnapshotter()
        self.payment_reader = payment_reader or PaymentInputSparseReader()
        self.payment_input = payment_input or PaymentInputWorkbook(reader=self.payment_reader)
        self.progress_index_reader = progress_index_reader or ActivityProgressIndexReader()
        self.position_engine = position_engine or PaymentPositionEngine()
        self.line_renderer = line_renderer or PaymentLineRenderer()

    def validate_workbook(self, workbook: Path) -> PaymentWorkbookValidation:
        return self.snapshotter.validate(Path(workbook))

    def create_payment_snapshot(self, source: Path, output: Path) -> PaymentSnapshotResult:
        return self.snapshotter.create_snapshot(Path(source), Path(output))

    def create_fake_payment_input(self, source: Path, output: Path, periods: int) -> PaymentInputResult:
        return self.payment_input.create(Path(source), Path(output), periods)

    def validate_payment_input(self, payment_workbook: Path, progress_workbook: Path | None = None) -> PaymentInputValidation:
        progress = Path(progress_workbook) if progress_workbook is not None else None
        return self.payment_input.validate(Path(payment_workbook), progress)

    def read_payment_requirements(self, payment_workbook: Path) -> PaymentInputData:
        return self.payment_reader.read(Path(payment_workbook))


    def prepare_payment_input(self, progress_workbook: Path, payment_workbook: Path) -> PaymentPreparationResult:
        """Validate and resolve the uploaded Payment Input with one read per workbook."""
        payment = self.payment_reader.read(Path(payment_workbook))
        progress = self.progress_index_reader.read(Path(progress_workbook))
        progress_ids = set(progress.activities)
        matched = sum(1 for activity_id in payment.activity_ids if activity_id in progress_ids)
        validation = PaymentInputValidation(
            workbook=payment.workbook,
            payment_sheet=payment.sheet,
            payment_periods=len(payment.periods),
            activity_rows=len(payment.activity_ids),
            matched_activities=matched,
            missing_activities=len(payment.activity_ids) - matched,
            populated_requirements=payment.populated_requirements,
        )
        positions = self.position_engine.resolve(payment, progress)
        return PaymentPreparationResult(validation=validation, positions=positions)

    def prepare_payment_positions(self, progress_workbook: Path, payment_workbook: Path) -> PaymentPositionResult:
        payment = self.payment_reader.read(Path(payment_workbook))
        progress = self.progress_index_reader.read(Path(progress_workbook))
        return self.position_engine.resolve(payment, progress)



    def render_payment_backbones(
        self,
        progress_workbook: Path,
        payment_workbook: Path,
        output_workbook: Path,
        period_ids: tuple[str, ...] = ("P01", "P02", "P03"),
    ) -> PaymentMultiLineRenderResult:
        """Render selected periods at their planned eligible boundaries (latest requirement)."""
        prepared = self.prepare_payment_input(Path(progress_workbook), Path(payment_workbook))
        by_id = {period.period_id: period for period in prepared.positions.periods}
        selected = []
        for period_id in period_ids:
            period = by_id.get(period_id)
            if period is None:
                from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
                raise PaymentWorkbookError(f"Payment period {period_id} was not found in Payment Input.")
            if period.points:
                selected.append(period)
        if not selected:
            from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
            raise PaymentWorkbookError("P01-P03 have no resolved requirements to render.")
        return self.line_renderer.render_periods(
            Path(progress_workbook), Path(output_workbook), tuple(selected)
        )

    def render_single_payment_line(
        self,
        progress_workbook: Path,
        payment_workbook: Path,
        output_workbook: Path,
        period_id: str = "P01",
    ) -> PaymentLineRenderResult:
        """Resolve and render one payment period at its planned eligible boundary."""
        prepared = self.prepare_payment_input(Path(progress_workbook), Path(payment_workbook))
        period = next((item for item in prepared.positions.periods if item.period_id == period_id), None)
        if period is None:
            from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
            raise PaymentWorkbookError(f"Payment period {period_id} was not found in Payment Input.")
        if not period.points:
            from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
            raise PaymentWorkbookError(f"{period_id} has no resolved requirements to render.")
        return self.line_renderer.render_single_period(
            Path(progress_workbook), Path(output_workbook), period
        )
