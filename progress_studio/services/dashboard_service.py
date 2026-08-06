from pathlib import Path

from progress_studio.infrastructure.excel.dashboard_workbook import build_dashboard_file


class DashboardService:
    def build(self, input_file: Path, output_file: Path, *, project_name: str | None = None) -> None:
        build_dashboard_file(input_file, output_file, project_name=project_name)
