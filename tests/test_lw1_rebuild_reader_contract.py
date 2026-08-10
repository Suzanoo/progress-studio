
from __future__ import annotations

import ast
from pathlib import Path

import pytest

from progress_studio.infrastructure.excel.rebuild_workbook_reader import (
    RebuildWorkbookReadError,
    RebuildWorkbookReader,
)


ROOT = Path(__file__).resolve().parents[1]
READER = ROOT / "progress_studio/infrastructure/excel/rebuild_workbook_reader.py"


def test_lw1_reader_boundary_does_not_import_openpyxl() -> None:
    """Analyze/probe must remain a sparse OOXML boundary, never a full workbook load."""
    tree = ast.parse(READER.read_text(encoding="utf-8"))
    imports = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imports.extend(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            imports.append(node.module or "")
    assert not any(name == "openpyxl" or name.startswith("openpyxl.") for name in imports)


def test_lw1_probe_contract_exposes_io_metrics(tmp_path: Path) -> None:
    """Probe reports cheap boundary metrics without retaining workbook cell objects."""
    # A non-XLSX file must fail through the reader's stable domain error.
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not an xlsx")
    with pytest.raises(RebuildWorkbookReadError):
        RebuildWorkbookReader().probe(bad)

    fields = set(RebuildWorkbookReader.__module__ for _ in [0])
    source = READER.read_text(encoding="utf-8")
    assert "package_bytes: int" in source
    assert "main_xml_bytes: int" in source
    assert "path.stat().st_size" in source
    assert "len(main_xml)" in source


def test_lw1_reader_contract_reads_only_workbook_metadata_shared_strings_and_main() -> None:
    source = READER.read_text(encoding="utf-8")
    assert 'package.read("xl/workbook.xml")' in source
    assert '"xl/sharedStrings.xml"' in source
    assert "package.read(sheet_map[self.MAIN_SHEET])" in source
    assert "styles.xml" not in source
    assert "drawings/" not in source
    assert "load_workbook" not in source


def test_lw1_probe_is_metadata_only_contract() -> None:
    """Probe remains cheap even after LW-2 adds the separate full MainDataset path."""
    source = READER.read_text(encoding="utf-8")
    assert "class RebuildWorkbookProbe" in source
    assert "activity_count: int" in source
    probe_block = source[source.index("    def probe("):source.index("    def read_main_dataset(")]
    assert "read_main_dataset" not in probe_block
    assert "_parse_main_dataset" not in probe_block
    assert "load_workbook" not in probe_block
