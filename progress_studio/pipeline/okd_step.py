from pathlib import Path
import shutil

from progress_studio.app.context import PipelineContext
from progress_studio.services.okd_service import OkdService


class OkdStep:
    name = "build-okd-sheets"

    def __init__(self, service: OkdService) -> None:
        self.service = service

    def execute(self, context: PipelineContext) -> PipelineContext:
        if context.distribution_workbook is None or context.project_folder is None:
            raise RuntimeError("DistributionStep must run before OkdStep.")

        output = context.distribution_workbook
        context.metadata["okd"] = self.service.build(output, output)
        context.okd_workbook = output
        context.output_workbook = output

        if context.amount_workbook is not None:
            mapping_copy = (
                context.project_folder
                / f"{context.source_xml.stem}_amount_mapping.xlsx"
            )
            shutil.copy2(context.amount_workbook, mapping_copy)
            context.metadata["amount_mapping_copy"] = mapping_copy

        return context
