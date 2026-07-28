from pathlib import Path
import tempfile
import unittest

from progress_studio.app.context import PipelineContext
from progress_studio.app.desktop import DesktopRunOptions, DesktopRunner
from progress_studio.app.pipeline import Pipeline, PipelineEvent


class FakeStep:
    def __init__(self, name: str) -> None:
        self.name = name

    def execute(self, context: PipelineContext) -> PipelineContext:
        context.metadata[self.name] = True
        return context


class DesktopPhase2Tests(unittest.TestCase):
    def test_pipeline_reports_step_events(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            xml = Path(folder) / "project.xml"
            xml.write_text("<Project />", encoding="utf-8")
            events: list[PipelineEvent] = []
            pipeline = Pipeline([FakeStep("one"), FakeStep("two")])
            context = PipelineContext(xml, "5", 1000)
            result = pipeline.run(context, observer=events.append)

        self.assertTrue(result.metadata["one"])
        self.assertEqual([event.status for event in events], ["started", "completed", "started", "completed"])
        self.assertEqual(events[-1].progress_percent, 100.0)

    def test_desktop_runner_uses_existing_pipeline_contract(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            xml = Path(folder) / "project.xml"
            xml.write_text("<Project />", encoding="utf-8")
            runner = DesktopRunner(lambda method: Pipeline([FakeStep(method)]))
            result = runner.run(DesktopRunOptions(xml, "5", 1000, "flat"))
        self.assertTrue(result.metadata["flat"])

    def test_desktop_runner_rejects_non_xml_input(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "project.xlsx"
            source.touch()
            runner = DesktopRunner(lambda method: Pipeline([]))
            with self.assertRaisesRegex(ValueError, "Primavera XML"):
                runner.run(DesktopRunOptions(source, "5", 1000))


if __name__ == "__main__":
    unittest.main()
