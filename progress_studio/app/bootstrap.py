from progress_studio.config import SETTINGS
from progress_studio.infrastructure.excel import ImportWorkbookWriter
from progress_studio.infrastructure.schedule_xml import NormalizedScheduleXmlReader
from progress_studio.pipeline.import_step import ImportStep
from progress_studio.pipeline.amount_step import AmountStep
from progress_studio.pipeline.progress_step import ProgressStep
from progress_studio.pipeline.distribution_step import DistributionStep
from progress_studio.pipeline.okd_step import OkdStep
from progress_studio.pipeline.monthly_main_step import MonthlyMainStep
from progress_studio.pipeline.schedule_step import ScheduleStep
from progress_studio.pipeline.timescale_step import TimescaleStep
from progress_studio.presentation.cli import CommandLineInterface
from progress_studio.presentation.distribution_prompt import DistributionPrompt
from progress_studio.services.import_service import ImportService
from progress_studio.services.amount_service import AmountService
from progress_studio.services.progress_service import ProgressService
from progress_studio.services.distribution_service import DistributionService
from progress_studio.services.schedule_service import ScheduleService
from progress_studio.services.schedule_workbook_service import ScheduleWorkbookService
from progress_studio.services.timescale_service import TimescaleService
from progress_studio.services.okd_service import OkdService
from progress_studio.services.monthly_main_service import MonthlyMainService

from .application import ProgressStudioApplication
from .pipeline import Pipeline


def build_application() -> ProgressStudioApplication:
    schedule_service = ScheduleService()
    import_service = ImportService(NormalizedScheduleXmlReader(), schedule_service, ImportWorkbookWriter())
    pipeline = Pipeline([
        ImportStep(import_service),
        ScheduleStep(ScheduleWorkbookService()),
        TimescaleStep(TimescaleService()),
        AmountStep(AmountService()),
        ProgressStep(ProgressService()),
        DistributionStep(DistributionService(), DistributionPrompt()),
        OkdStep(OkdService()),
        MonthlyMainStep(MonthlyMainService()),
    ])
    return ProgressStudioApplication(pipeline, CommandLineInterface(SETTINGS), SETTINGS)
