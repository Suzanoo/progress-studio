from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class PipelineContext:
    source_xml: Path
    cutoff_day: str
    amount_per_activity: float
    project_folder: Path | None = None
    working_folder: Path | None = None
    imported_workbook: Path | None = None
    scheduled_workbook: Path | None = None
    timescale_workbook: Path | None = None
    amount_workbook: Path | None = None
    progress_workbook: Path | None = None
    distribution_workbook: Path | None = None
    okd_workbook: Path | None = None
    output_workbook: Path | None = None
    metadata: dict[str, Any] = field(default_factory=dict)
