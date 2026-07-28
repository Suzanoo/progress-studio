from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from progress_studio.app.context import PipelineContext
from progress_studio.app.pipeline import Pipeline, PipelineEvent
from progress_studio.config import SETTINGS
from progress_studio.infrastructure.excel import ImportWorkbookWriter
from progress_studio.infrastructure.primavera import PrimaveraXmlReader
from progress_studio.pipeline.amount_step import AmountStep
from progress_studio.pipeline.distribution_step import DistributionStep
from progress_studio.pipeline.import_step import ImportStep
from progress_studio.pipeline.okd_step import OkdStep
from progress_studio.pipeline.progress_step import ProgressStep
from progress_studio.pipeline.schedule_step import ScheduleStep
from progress_studio.pipeline.timescale_step import TimescaleStep
from progress_studio.presentation.gui.distribution import FixedDistributionPrompt
from progress_studio.services.amount_service import AmountService
from progress_studio.services.distribution_service import DistributionService
from progress_studio.services.import_service import ImportService
from progress_studio.services.okd_service import OkdService
from progress_studio.services.progress_service import ProgressService
from progress_studio.services.schedule_service import ScheduleService
from progress_studio.services.schedule_workbook_service import ScheduleWorkbookService
from progress_studio.services.timescale_service import TimescaleService


@dataclass(frozen=True)
class DesktopRunOptions:
    source_xml: Path
    cutoff_day: str
    amount_per_activity: float
    distribution_method: str = "auto"


class DesktopRunner:
    """Application service that runs the existing pipeline for the GUI."""

    def __init__(self, pipeline_factory: Callable[[str], Pipeline] | None = None) -> None:
        self._pipeline_factory = pipeline_factory or build_desktop_pipeline

    def run(
        self,
        options: DesktopRunOptions,
        observer: Callable[[PipelineEvent], None] | None = None,
    ) -> PipelineContext:
        source = options.source_xml.expanduser().resolve()
        if not source.is_file():
            raise FileNotFoundError(f"Primavera XML file not found: {source}")
        if source.suffix.lower() != ".xml":
            raise ValueError("Input file must be a Primavera XML file.")
        if options.cutoff_day not in {"1", "2", "3", "4", "5", "6", "7"}:
            raise ValueError("Cutoff day must be between 1 and 7.")
        if options.amount_per_activity <= 0:
            raise ValueError("Placeholder amount must be greater than 0.")

        context = PipelineContext(
            source_xml=source,
            cutoff_day=options.cutoff_day,
            amount_per_activity=options.amount_per_activity,
        )
        return self._pipeline_factory(options.distribution_method).run(context, observer=observer)


def build_desktop_pipeline(distribution_method: str = "auto") -> Pipeline:
    schedule_service = ScheduleService()
    import_service = ImportService(
        PrimaveraXmlReader(), schedule_service, ImportWorkbookWriter()
    )
    return Pipeline(
        [
            ImportStep(import_service),
            ScheduleStep(ScheduleWorkbookService()),
            TimescaleStep(TimescaleService()),
            AmountStep(AmountService()),
            ProgressStep(ProgressService()),
            DistributionStep(
                DistributionService(), FixedDistributionPrompt(distribution_method)
            ),
            OkdStep(OkdService()),
        ]
    )
