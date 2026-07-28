from __future__ import annotations

from progress_studio.app.context import PipelineContext
from progress_studio.services.schedule_workbook_service import ScheduleWorkbookService


class ScheduleStep:
    name = "prepare-plan-actual-schedule"

    def __init__(self, service: ScheduleWorkbookService) -> None:
        self._service = service

    def execute(self, context: PipelineContext) -> PipelineContext:
        if context.imported_workbook is None or context.working_folder is None:
            raise RuntimeError("ImportStep must run before ScheduleStep.")
        output_folder = context.working_folder / "02_schedule"
        output_file = output_folder / f"{context.source_xml.stem}_schedule.xlsx"
        print("\n" + "=" * 72)
        print("[2/8] Prepare Plan / Actual Schedule")
        print("=" * 72)
        print(f"INPUT  : {context.imported_workbook}")
        print(f"OUTPUT : {output_file}")
        wbs_count, activity_count, final_rows = self._service.prepare(context.imported_workbook, output_file)
        context.scheduled_workbook = output_file
        context.metadata.update(schedule_wbs_count=wbs_count, schedule_activity_count=activity_count, schedule_rows=final_rows)
        print(f"WBS        : {wbs_count:,}")
        print(f"ACTIVITIES : {activity_count:,}")
        print(f"ROWS       : {final_rows:,}")
        print(f"OK         : {output_file}")
        return context
