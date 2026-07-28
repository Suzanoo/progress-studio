from pathlib import Path
import shutil

from progress_studio.app.context import PipelineContext
from progress_studio.presentation.distribution_prompt import DistributionPrompt
from progress_studio.services.distribution_service import DistributionService


class DistributionStep:
    name = "generate-plan-distribution"

    def __init__(self, service: DistributionService, prompt: DistributionPrompt) -> None:
        self.service = service
        self.prompt = prompt

    def execute(self, context: PipelineContext) -> PipelineContext:
        if context.progress_workbook is None or context.project_folder is None:
            raise RuntimeError("ProgressStep must run before DistributionStep.")
        last_file: Path | None = None
        while True:
            method = self.prompt.choose_method()
            if method is None:
                if last_file is None:
                    raise RuntimeError("Cancelled before Plan Distribution was generated.")
                context.distribution_workbook = last_file
                return context
            candidate = context.project_folder / f"{context.source_xml.stem}_{method}_progress.xlsx"
            result = self.service.generate(context.progress_workbook, candidate, method=method, debug=True)
            context.metadata["distribution"] = result
            last_file = candidate
            action = self.prompt.review(candidate, method)
            if action == "retry":
                continue
            if action == "accept":
                final_file = context.project_folder / f"{context.source_xml.stem}_final_progress.xlsx"
                shutil.copy2(candidate, final_file)
                context.distribution_workbook = final_file
            else:
                context.distribution_workbook = candidate
            return context
