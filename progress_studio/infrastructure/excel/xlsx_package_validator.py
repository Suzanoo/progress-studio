from __future__ import annotations

from pathlib import Path
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class WorkbookPackageValidationError(ValueError):
    """Raised when an exported XLSX package is structurally inconsistent."""


def validate_xlsx_tables(path: Path) -> None:
    """Validate table XML and reject duplicate worksheet/table AutoFilters.

    openpyxl can reopen packages that desktop Excel later repairs.  This check
    inspects the OOXML package directly before the temporary export replaces
    the user's destination file.
    """
    path = Path(path)
    try:
        with ZipFile(path) as archive:
            names = set(archive.namelist())
            table_names: set[str] = set()
            for table_part in sorted(name for name in names if name.startswith("xl/tables/table") and name.endswith(".xml")):
                root = ET.fromstring(archive.read(table_part))
                ref = root.attrib.get("ref", "")
                display_name = root.attrib.get("displayName", "")
                if not ref:
                    raise WorkbookPackageValidationError(f"Table has no range: {table_part}")
                if not display_name:
                    raise WorkbookPackageValidationError(f"Table has no displayName: {table_part}")
                if display_name in table_names:
                    raise WorkbookPackageValidationError(f"Duplicate table name: {display_name}")
                table_names.add(display_name)

                columns = root.find(f"{{{_MAIN_NS}}}tableColumns")
                if columns is None:
                    raise WorkbookPackageValidationError(f"Table has no columns: {table_part}")
                declared = int(columns.attrib.get("count", "0"))
                actual = len(columns.findall(f"{{{_MAIN_NS}}}tableColumn"))
                if declared != actual:
                    raise WorkbookPackageValidationError(
                        f"Table column count mismatch in {table_part}: declared {declared}, found {actual}"
                    )

            for sheet_part in sorted(name for name in names if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")):
                root = ET.fromstring(archive.read(sheet_part))
                table_parts = root.find(f"{{{_MAIN_NS}}}tableParts")
                if table_parts is None:
                    continue
                if root.find(f"{{{_MAIN_NS}}}autoFilter") is not None:
                    raise WorkbookPackageValidationError(
                        f"Worksheet AutoFilter overlaps an Excel Table in {sheet_part}"
                    )

                rels_part = sheet_part.rsplit("/", 1)[0] + "/_rels/" + sheet_part.rsplit("/", 1)[1] + ".rels"
                if rels_part not in names:
                    raise WorkbookPackageValidationError(f"Missing table relationships for {sheet_part}")
                rels_root = ET.fromstring(archive.read(rels_part))
                relationships = {
                    rel.attrib.get("Id"): rel.attrib.get("Target", "")
                    for rel in rels_root.findall(f"{{{_REL_NS}}}Relationship")
                }
                for table_part in table_parts.findall(f"{{{_MAIN_NS}}}tablePart"):
                    rel_id = table_part.attrib.get(f"{{{_DOC_REL_NS}}}id")
                    target = relationships.get(rel_id)
                    if not target:
                        raise WorkbookPackageValidationError(
                            f"Missing relationship {rel_id!r} for {sheet_part}"
                        )
                    normalized = "xl/worksheets/" + target if not target.startswith("/") else target.lstrip("/")
                    normalized = str(Path(normalized)).replace("\\", "/")
                    while "/../" in normalized:
                        before, after = normalized.split("/../", 1)
                        normalized = before.rsplit("/", 1)[0] + "/" + after
                    if normalized not in names:
                        raise WorkbookPackageValidationError(
                            f"Missing table target {normalized} referenced by {sheet_part}"
                        )
    except (BadZipFile, ET.ParseError) as exc:
        raise WorkbookPackageValidationError(f"Invalid XLSX package: {exc}") from exc
