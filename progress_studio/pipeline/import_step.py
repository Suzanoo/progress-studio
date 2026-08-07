from __future__ import annotations

from datetime import datetime

from progress_studio.app.context import PipelineContext
from progress_studio.config import SETTINGS
from progress_studio.infrastructure.filesystem import desktop_path
from progress_studio.services.import_service import ImportService


class ImportStep:
    name = "import-schedule-xml"

    def __init__(self, service: ImportService) -> None:
        self._service = service

    def execute(self, context: PipelineContext) -> PipelineContext:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        project_name = context.source_xml.stem
        context.project_folder = desktop_path() / SETTINGS.output_root_name / f"{project_name}_{timestamp}"
        context.working_folder = context.project_folder / "working_files"
        output_folder = context.working_folder / "01_import"
        output_file = output_folder / f"{project_name}_imported.xlsx"
        print("\n" + "=" * 72)
        print("Import Schedule XML")
        print("=" * 72)
        print(f"INPUT  : {context.source_xml}")
        print(f"OUTPUT : {output_file}")
        name, wbs_count, activity_count = self._service.import_xml(context.source_xml, output_file)
        context.imported_workbook = output_file
        context.metadata.update(project_name=name, wbs_count=wbs_count, activity_count=activity_count)
        print(f"WBS        : {wbs_count:,}")
        print(f"ACTIVITIES : {activity_count:,}")
        print(f"OK         : {output_file}")
        return context
