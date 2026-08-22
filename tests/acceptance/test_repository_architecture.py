import importlib
import tempfile
import unittest
from pathlib import Path

from tests._paths import REPO_ROOT

from progress_studio.app.context import PipelineContext
from tests.acceptance.test_release_contract import run_example

ROOT = REPO_ROOT


class Ms8ArchitectureCleanupTests(unittest.TestCase):
    def test_obsolete_root_modules_are_removed(self):
        self.assertFalse((ROOT / "excel_toolkit").exists())
        self.assertFalse((ROOT / "distribution").exists())
        self.assertFalse((ROOT / "excel_theme.py").exists())

    def test_new_owners_exist(self):
        self.assertTrue((ROOT / "progress_studio/domain/activity_id.py").is_file())
        self.assertTrue((ROOT / "progress_studio/services/distribution/curves.py").is_file())
        self.assertTrue((ROOT / "progress_studio/services/distribution/auto.py").is_file())
        self.assertTrue((ROOT / "progress_studio/services/distribution/distribution_rules.json").is_file())
        self.assertTrue((ROOT / "progress_studio/infrastructure/excel/styles.py").is_file())

    def test_no_import_references_removed_modules(self):
        forbidden = (
            "excel_toolkit",
            "from distribution",
            "import distribution",
            "from excel_theme",
            "import excel_theme",
        )
        for path in (ROOT / "progress_studio").rglob("*.py"):
            text = path.read_text(encoding="utf-8")
            for phrase in forbidden:
                self.assertNotIn(phrase, text, str(path))

    def test_every_application_module_imports(self):
        for path in (ROOT / "progress_studio").rglob("*.py"):
            relative = path.relative_to(ROOT).with_suffix("")
            module = ".".join(relative.parts)
            importlib.import_module(module)

    def test_end_to_end_manifest_is_unchanged(self):
        with tempfile.TemporaryDirectory() as td:
            context: PipelineContext = run_example(Path(td))
            self.assertEqual(context.metadata["activity_count"], 172)
            self.assertEqual(context.metadata["wbs_count"], 82)
            self.assertEqual(context.metadata["progress"].weekly_columns, 76)
            self.assertEqual(context.metadata["okd"].table_rows, 510)
            self.assertEqual(context.metadata["okd"].checked_links, 40290)


if __name__ == "__main__":
    unittest.main()
