from progress_studio.app.context import PipelineContext
from progress_studio.services.monthly_main_service import MonthlyMainService


class MonthlyMainStep:
    name = "build-monthly-main-view"

    def __init__(self, service: MonthlyMainService) -> None:
        self.service = service

    def execute(self, context: PipelineContext) -> PipelineContext:
        if context.okd_workbook is None:
            raise RuntimeError("OkdStep must run before MonthlyMainStep.")

        output = context.okd_workbook
        context.metadata["monthly_main"] = self.service.build(output, output)
        context.monthly_workbook = output
        context.output_workbook = output
        return context
