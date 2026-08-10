from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from progress_studio.domain.main_dataset import MainDataset, MainPeriod, MainRow


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class RebuildWorkbookReadError(ValueError):
    pass


@dataclass(frozen=True, slots=True)
class RebuildWorkbookProbe:
    workbook: Path
    sheet_names: tuple[str, ...]
    main_rows: int
    main_columns: int
    activity_count: int
    package_bytes: int
    main_xml_bytes: int


class RebuildWorkbookReader:
    """Sparse XLSX/XLSM reader for the standalone rebuild boundary.

    Reads workbook metadata plus `main` worksheet XML only. Generated sheets,
    dashboards, formulas, drawings, and styles are never loaded into Python.
    """

    MAIN_SHEET = "main"

    def probe(self, workbook_path: Path) -> RebuildWorkbookProbe:
        path = Path(workbook_path)
        try:
            with ZipFile(path, "r") as package:
                sheet_map = self._sheet_map(package)
                if self.MAIN_SHEET not in sheet_map:
                    raise RebuildWorkbookReadError(
                        "Rebuild requires worksheet 'main'."
                    )
                shared = self._shared_strings(package)
                main_xml = package.read(sheet_map[self.MAIN_SHEET])
                main_root = ET.fromstring(main_xml)
                rows, cols, activities = self._main_metrics(main_root, shared)
                return RebuildWorkbookProbe(
                    workbook=path,
                    sheet_names=tuple(sheet_map.keys()),
                    main_rows=rows,
                    main_columns=cols,
                    activity_count=activities,
                    package_bytes=path.stat().st_size,
                    main_xml_bytes=len(main_xml),
                )
        except RebuildWorkbookReadError:
            raise
        except Exception as exc:
            raise RebuildWorkbookReadError(
                f"Workbook structure could not be read: {exc}"
            ) from exc


    def read_main_dataset(self, workbook_path: Path) -> MainDataset:
        """Parse `main` once into an immutable domain dataset without openpyxl."""
        path = Path(workbook_path)
        try:
            with ZipFile(path, "r") as package:
                sheet_map = self._sheet_map(package)
                if self.MAIN_SHEET not in sheet_map:
                    raise RebuildWorkbookReadError("Rebuild requires worksheet 'main'.")
                shared = self._shared_strings(package)
                root = ET.fromstring(package.read(sheet_map[self.MAIN_SHEET]))
                return self._parse_main_dataset(path, root, shared)
        except RebuildWorkbookReadError:
            raise
        except Exception as exc:
            raise RebuildWorkbookReadError(f"Workbook structure could not be read: {exc}") from exc

    @classmethod
    def _parse_main_dataset(cls, path: Path, root: ET.Element, shared_strings: list[str]) -> MainDataset:
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise RebuildWorkbookReadError("Worksheet 'main' has no data.")
        rows = sheet_data.findall(f"{{{MAIN_NS}}}row")
        header_row, headers = cls._find_main_header(rows, shared_strings)

        # Weekly grammar is intentionally structural: Wn sits one row above the
        # reporting date header. This avoids loading workbook style metadata merely to infer dates.
        by_number = {int(r.attrib.get("r", "0")): r for r in rows}
        week_row = by_number.get(header_row - 1)
        header_xml_row = by_number.get(header_row)
        periods: list[MainPeriod] = []
        if week_row is not None and header_xml_row is not None:
            week_cells = cls._row_cells(week_row)
            header_cells = cls._row_cells(header_xml_row)
            for col, cell in sorted(week_cells.items()):
                key = cls._cell_text(cell, shared_strings).strip()
                if re.fullmatch(r"W\d+", key, flags=re.IGNORECASE):
                    raw = cls._cell_value(header_cells.get(col), shared_strings)
                    periods.append(MainPeriod(col, key, cls._coerce_date(raw)))

        parsed_rows: list[MainRow] = []
        for row in rows:
            row_num = int(row.attrib.get("r", "0"))
            if row_num <= header_row:
                continue
            cells = cls._row_cells(row)
            def val(name: str):
                col = headers.get(name)
                return cls._cell_value(cells.get(col), shared_strings) if col else None
            row_type = str(val("row type") or "").strip()
            pa = str(val("p/a") or "").strip()
            wbs = str(val("wbs") or "").strip()
            description = str(val("description") or "").strip()
            activity_id = str(val("activity id") or "").strip()
            period_values = tuple(
                (p.column, cls._coerce_float(cls._cell_value(cells.get(p.column), shared_strings)))
                for p in periods
            )
            parsed_rows.append(MainRow(
                row_number=row_num,
                row_type=row_type,
                pa=pa,
                wbs=wbs,
                description=description,
                activity_id=activity_id,
                outline_level=cls._coerce_int(val("outline level")),
                plan_start=cls._coerce_date(val("plan start")),
                plan_finish=cls._coerce_date(val("plan finish")),
                amount=cls._coerce_float(val("amount")),
                percent_complete=cls._coerce_float(val("% complete")),
                period_values=period_values,
            ))
        return MainDataset(
            workbook_name=path.name,
            header_row=header_row,
            headers=tuple(sorted(headers.items(), key=lambda item: item[1])),
            periods=tuple(periods),
            rows=tuple(parsed_rows),
        )

    @classmethod
    def _find_main_header(cls, rows: list[ET.Element], shared_strings: list[str]) -> tuple[int, dict[str, int]]:
        required = {"row type", "p/a", "activity id"}
        for row in rows:
            row_num = int(row.attrib.get("r", "0"))
            if row_num > 30:
                break
            found: dict[str, int] = {}
            for col, cell in cls._row_cells(row).items():
                text = cls._cell_text(cell, shared_strings).strip().lower()
                if text:
                    found[text] = col
            if required.issubset(found):
                return row_num, found
        raise RebuildWorkbookReadError(
            "Worksheet 'main' is missing required headers: Row Type, P/A, Activity ID."
        )

    @staticmethod
    def _coerce_float(value: object) -> float | None:
        if value in (None, ""):
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _coerce_int(value: object) -> int | None:
        number = RebuildWorkbookReader._coerce_float(value)
        return int(number) if number is not None else None

    @staticmethod
    def _coerce_date(value: object) -> datetime | None:
        if value in (None, ""):
            return None
        if isinstance(value, datetime):
            return value
        if isinstance(value, (int, float)):
            return datetime(1899, 12, 30) + timedelta(days=float(value))
        text = str(value).strip()
        try:
            return datetime.fromisoformat(text)
        except ValueError:
            try:
                return datetime(1899, 12, 30) + timedelta(days=float(text))
            except ValueError:
                return None


    @classmethod
    def _sheet_map(cls, package: ZipFile) -> dict[str, str]:
        workbook_root = ET.fromstring(package.read("xl/workbook.xml"))
        rel_root = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        rels = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rel_root.findall(f"{{{PKG_REL_NS}}}Relationship")
        }

        sheets = workbook_root.find(f"{{{MAIN_NS}}}sheets")
        if sheets is None:
            return {}

        result: dict[str, str] = {}
        for sheet in sheets.findall(f"{{{MAIN_NS}}}sheet"):
            name = sheet.attrib.get("name", "")
            rel_id = sheet.attrib.get(f"{{{REL_NS}}}id")
            target = rels.get(rel_id or "", "")
            if not name or not target:
                continue
            target = target.replace("\\", "/")
            if target.startswith("/"):
                path = target.lstrip("/")
            elif target.startswith("xl/"):
                path = target
            else:
                # worksheet relationships are normally relative to /xl
                parts = []
                for part in ("xl/" + target).split("/"):
                    if part == "..":
                        if parts:
                            parts.pop()
                    elif part not in ("", "."):
                        parts.append(part)
                path = "/".join(parts)
            result[name] = path
        return result

    @staticmethod
    def _shared_strings(package: ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in package.namelist():
            return []
        root = ET.fromstring(package.read("xl/sharedStrings.xml"))
        result: list[str] = []
        for item in root.findall(f"{{{MAIN_NS}}}si"):
            result.append(
                "".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t"))
            )
        return result

    @classmethod
    def _main_metrics(
        cls,
        root: ET.Element,
        shared_strings: list[str],
    ) -> tuple[int, int, int]:
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise RebuildWorkbookReadError("Worksheet 'main' has no data.")

        max_row = 0
        max_col = 0
        header_row = None
        columns: dict[str, int] = {}
        rows = sheet_data.findall(f"{{{MAIN_NS}}}row")

        for row in rows:
            row_num = int(row.attrib.get("r", "0"))
            max_row = max(max_row, row_num)
            cells = cls._row_cells(row)
            if cells:
                max_col = max(max_col, max(cells))
            if header_row is None and row_num <= 30:
                found: dict[str, int] = {}
                for col, cell in cells.items():
                    text = cls._cell_text(cell, shared_strings).strip().lower()
                    if text in {"row type", "p/a", "activity id"}:
                        found[text] = col
                if {"row type", "p/a", "activity id"}.issubset(found):
                    header_row = row_num
                    columns = found

        if header_row is None:
            raise RebuildWorkbookReadError(
                "Worksheet 'main' is missing required headers: "
                "Row Type, P/A, Activity ID."
            )

        activity_count = 0
        for row in rows:
            row_num = int(row.attrib.get("r", "0"))
            if row_num <= header_row:
                continue
            cells = cls._row_cells(row)
            row_type = cls._cell_text(
                cells.get(columns["row type"]), shared_strings
            ).strip().lower()
            pa = cls._cell_text(
                cells.get(columns["p/a"]), shared_strings
            ).strip().upper()
            activity_id = cls._cell_text(
                cells.get(columns["activity id"]), shared_strings
            ).strip()
            if row_type == "activity" and pa == "P" and activity_id:
                activity_count += 1

        return max_row, max_col, activity_count

    @staticmethod
    def _row_cells(row: ET.Element) -> dict[int, ET.Element]:
        result: dict[int, ET.Element] = {}
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            match = re.match(r"([A-Z]+)", cell.attrib.get("r", ""))
            if match:
                result[RebuildWorkbookReader._column_index(match.group(1))] = cell
        return result

    @staticmethod
    def _column_index(letters: str) -> int:
        value = 0
        for ch in letters:
            value = value * 26 + ord(ch) - 64
        return value

    @staticmethod
    def _cell_value(cell: ET.Element | None, shared_strings: list[str]):
        if cell is None:
            return None
        text = RebuildWorkbookReader._cell_text(cell, shared_strings)
        if text == "":
            return None
        cell_type = cell.attrib.get("t")
        if cell_type in {"s", "inlineStr", "str"}:
            return text
        try:
            number = float(text)
            return int(number) if number.is_integer() else number
        except ValueError:
            return text

    @staticmethod
    def _cell_text(cell: ET.Element | None, shared_strings: list[str]) -> str:
        if cell is None:
            return ""
        cell_type = cell.attrib.get("t")
        if cell_type == "inlineStr":
            inline = cell.find(f"{{{MAIN_NS}}}is")
            if inline is None:
                return ""
            return "".join(
                node.text or "" for node in inline.iter(f"{{{MAIN_NS}}}t")
            )
        value = cell.find(f"{{{MAIN_NS}}}v")
        if value is None or value.text is None:
            return ""
        if cell_type == "s":
            try:
                return shared_strings[int(value.text)]
            except (ValueError, IndexError):
                return ""
        return value.text
