import zipfile

from progress_studio.infrastructure.excel.xlsx_package_preservation import (
    restore_opaque_workbook_parts,
)


def _write_zip(path, parts):
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _read(path, name):
    with zipfile.ZipFile(path, "r") as archive:
        return archive.read(name)


def test_pb4r_restores_existing_drawings_and_external_links_byte_for_byte(tmp_path):
    source = tmp_path / "source.xlsx"
    target = tmp_path / "target.xlsx"

    _write_zip(
        source,
        {
            "xl/workbook.xml": b"source-workbook",
            "xl/drawings/drawing1.xml": b"ORIGINAL-DRAWING",
            "xl/drawings/_rels/drawing1.xml.rels": b"ORIGINAL-DRAWING-RELS",
            "xl/externalLinks/externalLink1.xml": b"ORIGINAL-LINK",
            "xl/externalLinks/_rels/externalLink1.xml.rels": b"ORIGINAL-LINK-RELS",
        },
    )
    _write_zip(
        target,
        {
            "xl/workbook.xml": b"target-workbook",
            "xl/worksheets/sheet99.xml": b"NEW-PB-SHEET",
            "xl/drawings/drawing1.xml": b"REWRITTEN-DRAWING",
            "xl/drawings/_rels/drawing1.xml.rels": b"REWRITTEN-DRAWING-RELS",
            "xl/externalLinks/externalLink1.xml": b"REWRITTEN-LINK",
            "xl/externalLinks/_rels/externalLink1.xml.rels": b"REWRITTEN-LINK-RELS",
        },
    )

    restored = restore_opaque_workbook_parts(source, target)

    assert "xl/drawings/drawing1.xml" in restored
    assert "xl/externalLinks/externalLink1.xml" in restored
    assert _read(target, "xl/drawings/drawing1.xml") == b"ORIGINAL-DRAWING"
    assert _read(target, "xl/drawings/_rels/drawing1.xml.rels") == b"ORIGINAL-DRAWING-RELS"
    assert _read(target, "xl/externalLinks/externalLink1.xml") == b"ORIGINAL-LINK"
    assert _read(target, "xl/externalLinks/_rels/externalLink1.xml.rels") == b"ORIGINAL-LINK-RELS"
    assert _read(target, "xl/worksheets/sheet99.xml") == b"NEW-PB-SHEET"
    assert _read(target, "xl/workbook.xml") == b"target-workbook"
