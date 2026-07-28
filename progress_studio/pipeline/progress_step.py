from progress_studio.app.context import PipelineContext
from progress_studio.services.progress_service import ProgressService


class ProgressStep:
    name = "build-progress-workbook"

    def __init__(self, service: ProgressService) -> None:
        self.service = service

    def execute(self, context: PipelineContext) -> PipelineContext:
        if context.amount_workbook is None or context.working_folder is None:
            raise RuntimeError("AmountStep must run before ProgressStep.")
        output = context.working_folder / "05_progress_workbook.xlsx"
        context.metadata["progress"] = self.service.build(context.amount_workbook, output)
        context.progress_workbook = output
        return context
