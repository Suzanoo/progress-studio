from __future__ import annotations

from collections.abc import Sequence

from progress_studio.config import SETTINGS, Settings
from progress_studio.presentation.cli import CommandLineInterface

from .context import PipelineContext
from .pipeline import Pipeline


class ApplicationError(RuntimeError):
    """Raised when application input is invalid."""


class ProgressStudioApplication:
    def __init__(
        self,
        pipeline: Pipeline,
        cli: CommandLineInterface,
        settings: Settings = SETTINGS,
    ) -> None:
        self._pipeline = pipeline
        self._cli = cli
        self._settings = settings

    def run(self, argv: Sequence[str] | None = None) -> int:
        try:
            options = self._cli.parse(argv)
            source_xml = options.input_file or self._cli.select_xml_file()
            if source_xml is None:
                print("File selection cancelled.")
                return 0
            source_xml = source_xml.expanduser().resolve()
            if not source_xml.is_file():
                raise ApplicationError(f"Input XML file not found: {source_xml}")
            if options.amount <= 0:
                raise ApplicationError("--amount must be greater than 0.")

            cutoff_day = options.cutoff_day or self._cli.select_cutoff_day()
            context = PipelineContext(
                source_xml=source_xml,
                cutoff_day=cutoff_day,
                amount_per_activity=options.amount,
            )
            result = self._pipeline.run(context)
            if result.output_workbook:
                print(f"\nCompleted: {result.output_workbook}")
            return 0
        except KeyboardInterrupt:
            print("\nOperation cancelled.")
            return 130
        except Exception as exc:
            print(f"\nERROR\n{exc}")
            return 1
