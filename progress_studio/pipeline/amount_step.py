from __future__ import annotations

from progress_studio.app.context import PipelineContext
from progress_studio.services.amount_service import AmountService


class AmountStep:
    name = "build-and-apply-amounts"

    def __init__(self, service: AmountService) -> None:
        self.service = service

    def execute(self, context: PipelineContext) -> PipelineContext:
        if context.timescale_workbook is None or context.working_folder is None:
            raise RuntimeError("TimescaleStep must run before AmountStep.")
        mapping_result = self.service.build_mapping(
            context.timescale_workbook,
            context.amount_per_activity,
        )
        output = context.working_folder / "04_amount_mapped.xlsx"
        application_result = self.service.apply_mapping(context.timescale_workbook, output)
        context.amount_workbook = output
        context.metadata["amount_mapping"] = mapping_result
        context.metadata["amount_application"] = application_result
        return context
