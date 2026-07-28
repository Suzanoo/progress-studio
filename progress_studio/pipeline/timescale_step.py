from __future__ import annotations

from progress_studio.app.context import PipelineContext
from progress_studio.services.timescale_service import TimescaleService


class TimescaleStep:
    name = "build-timescale"

    def __init__(self, service: TimescaleService) -> None:
        self.service = service

    def execute(self, context: PipelineContext) -> PipelineContext:
        if context.scheduled_workbook is None or context.working_folder is None:
            raise RuntimeError("ScheduleStep must run before TimescaleStep.")
        output_folder = context.working_folder / "03_timescale"
        output_folder.mkdir(parents=True, exist_ok=True)
        output = output_folder / f"{context.scheduled_workbook.stem}_timescale.xlsx"
        context.timescale_workbook = self.service.build(
            context.scheduled_workbook, output, context.cutoff_day
        )
        return context
