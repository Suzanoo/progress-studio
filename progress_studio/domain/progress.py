from dataclasses import dataclass


@dataclass(frozen=True)
class ProgressBuildResult:
    activities_with_amount: int
    activities_without_amount: int
    wbs_rollups: int
    project_rollups: int
    weekly_columns: int
