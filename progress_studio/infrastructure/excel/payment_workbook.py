from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from progress_studio.domain.payment_models import PaymentSnapshotResult, PaymentWorkbookValidation


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)
ET.register_namespace("", PKG_REL_NS)


class PaymentWorkbookError(ValueError):
    pass


class PaymentWorkbookSnapshotter:
    """Fast package-level snapshot of exported ``main`` into ``Payment``.

    MS-PAY1 intentionally does not recalculate or render anything. Copying the
    worksheet XML keeps the Payment sheet visually identical to main while
    avoiding a full openpyxl load/save of a large exported workbook.
    """

    MAIN_SHEET = "main"
    PAYMENT_SHEET = "Payment"
    ACTIVITY_ID_HEADERS = {"activity id", "activity_id", "activityid", "act id", "act. id"}

    def validate(self, workbook_path: Path) -> PaymentWorkbookValidation:
        path = Path(workbook_path)
        if not path.exists():
            raise PaymentWorkbookError(f"Workbook was not found: {path}")
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise PaymentWorkbookError("Select a Progress Studio .xlsx or .xlsm workbook.")

        try:
            with ZipFile(path, "r") as package:
                sheet_map = self._sheet_map(package)
                if self.MAIN_SHEET not in sheet_map:
                    raise PaymentWorkbookError("Worksheet 'main' was not found. Select an exported Progress Studio workbook.")
                main_part = sheet_map[self.MAIN_SHEET]
                root = ET.fromstring(package.read(main_part))
                max_row, max_column = self._dimensions(root)
                shared_strings = self._shared_strings(package)
                activity_rows = self._count_activity_rows(root, shared_strings)
                if activity_rows <= 0:
                    raise PaymentWorkbookError("No activity rows were found in the 'main' worksheet.")
                return PaymentWorkbookValidation(path, self.MAIN_SHEET, activity_rows, max_row, max_column)
        except PaymentWorkbookError:
            raise
        except Exception as exc:
            raise PaymentWorkbookError(f"Workbook cannot be opened: {exc}") from exc

    def create_snapshot(self, source: Path, output: Path) -> PaymentSnapshotResult:
        validation = self.validate(source)
        source_path = Path(source)
        output_path = Path(output)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            with ZipFile(source_path, "r") as src:
                workbook_root = ET.fromstring(src.read("xl/workbook.xml"))
                rels_root = ET.fromstring(src.read("xl/_rels/workbook.xml.rels"))
                content_root = ET.fromstring(src.read("[Content_Types].xml"))

                sheets = workbook_root.find(f"{{{MAIN_NS}}}sheets")
                if sheets is None:
                    raise PaymentWorkbookError("Workbook sheet metadata is missing.")
                rel_map = {
                    rel.attrib["Id"]: rel.attrib["Target"]
                    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
                }

                main_node = next((s for s in sheets if s.attrib.get("name") == self.MAIN_SHEET), None)
                if main_node is None:
                    raise PaymentWorkbookError("Worksheet 'main' was not found.")
                main_rid = main_node.attrib[f"{{{DOC_REL_NS}}}id"]
                main_part = self._normalise_sheet_target(rel_map[main_rid])
                main_xml = src.read(main_part)

                payment_node = next((s for s in sheets if s.attrib.get("name") == self.PAYMENT_SHEET), None)
                replaced = payment_node is not None
                if payment_node is not None:
                    payment_rid = payment_node.attrib[f"{{{DOC_REL_NS}}}id"]
                    payment_part = self._normalise_sheet_target(rel_map[payment_rid])
                else:
                    payment_rid = self._next_relationship_id(rels_root)
                    payment_part = self._next_sheet_part(src)
                    sheet_id = str(max(int(s.attrib.get("sheetId", "0")) for s in sheets) + 1)
                    ET.SubElement(
                        sheets,
                        f"{{{MAIN_NS}}}sheet",
                        {
                            "name": self.PAYMENT_SHEET,
                            "sheetId": sheet_id,
                            f"{{{DOC_REL_NS}}}id": payment_rid,
                        },
                    )
                    ET.SubElement(
                        rels_root,
                        f"{{{PKG_REL_NS}}}Relationship",
                        {
                            "Id": payment_rid,
                            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                            "Target": str(PurePosixPath(payment_part).relative_to("xl")),
                        },
                    )
                    ET.SubElement(
                        content_root,
                        f"{{{CONTENT_NS}}}Override",
                        {
                            "PartName": "/" + payment_part,
                            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        },
                    )

                main_rels_part = self._worksheet_rels_part(main_part)
                payment_rels_part = self._worksheet_rels_part(payment_part)
                main_rels = src.read(main_rels_part) if main_rels_part in src.namelist() else None

                replacements: dict[str, bytes] = {
                    "xl/workbook.xml": ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True),
                    "xl/_rels/workbook.xml.rels": ET.tostring(rels_root, encoding="utf-8", xml_declaration=True),
                    "[Content_Types].xml": ET.tostring(content_root, encoding="utf-8", xml_declaration=True),
                    payment_part: main_xml,
                }
                if main_rels is not None:
                    replacements[payment_rels_part] = main_rels

                self._write_package(src, output_path, replacements, remove={payment_rels_part} if main_rels is None else set())
        except PaymentWorkbookError:
            raise
        except Exception as exc:
            raise PaymentWorkbookError(f"Payment snapshot could not be created: {exc}") from exc

        return PaymentSnapshotResult(source_path, output_path, self.PAYMENT_SHEET, replaced, validation.activity_rows)

    @staticmethod
    def _write_package(src: ZipFile, output: Path, replacements: dict[str, bytes], remove: set[str]) -> None:
        # Write beside the destination then atomically replace, so an interrupted
        # snapshot never leaves the requested output path half-written.
        with tempfile.NamedTemporaryFile(prefix="payment_", suffix=output.suffix, dir=output.parent, delete=False) as handle:
            temp_path = Path(handle.name)
        try:
            with ZipFile(temp_path, "w") as dst:
                existing = set(src.namelist())
                for info in src.infolist():
                    if info.filename in remove:
                        continue
                    data = replacements.get(info.filename, src.read(info.filename))
                    new_info = ZipInfo(info.filename, date_time=info.date_time)
                    new_info.compress_type = info.compress_type
                    new_info.comment = info.comment
                    new_info.extra = info.extra
                    new_info.internal_attr = info.internal_attr
                    new_info.external_attr = info.external_attr
                    new_info.create_system = info.create_system
                    dst.writestr(new_info, data)
                for name, data in replacements.items():
                    if name not in existing:
                        dst.writestr(name, data, compress_type=ZIP_DEFLATED)
            shutil.move(str(temp_path), str(output))
        finally:
            temp_path.unlink(missing_ok=True)

    @staticmethod
    def _sheet_map(package: ZipFile) -> dict[str, str]:
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        rels = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        sheets = workbook.find(f"{{{MAIN_NS}}}sheets")
        if sheets is None:
            return {}
        result = {}
        for sheet in sheets:
            rid = sheet.attrib.get(f"{{{DOC_REL_NS}}}id")
            if rid and rid in rel_map:
                result[sheet.attrib["name"]] = PaymentWorkbookSnapshotter._normalise_sheet_target(rel_map[rid])
        return result

    @staticmethod
    def _normalise_sheet_target(target: str) -> str:
        target = target.lstrip("/")
        return target if target.startswith("xl/") else "xl/" + target

    @staticmethod
    def _worksheet_rels_part(sheet_part: str) -> str:
        p = PurePosixPath(sheet_part)
        return str(p.parent / "_rels" / f"{p.name}.rels")

    @staticmethod
    def _next_relationship_id(rels_root: ET.Element) -> str:
        numbers = []
        for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
            match = re.fullmatch(r"rId(\d+)", rel.attrib.get("Id", ""))
            if match:
                numbers.append(int(match.group(1)))
        return f"rId{max(numbers, default=0) + 1}"

    @staticmethod
    def _next_sheet_part(package: ZipFile) -> str:
        numbers = []
        for name in package.namelist():
            match = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", name)
            if match:
                numbers.append(int(match.group(1)))
        return f"xl/worksheets/sheet{max(numbers, default=0) + 1}.xml"

    @staticmethod
    def _dimensions(root: ET.Element) -> tuple[int, int]:
        dimension = root.find(f"{{{MAIN_NS}}}dimension")
        ref = dimension.attrib.get("ref", "A1") if dimension is not None else "A1"
        end = ref.split(":")[-1]
        match = re.fullmatch(r"([A-Z]+)(\d+)", end)
        if not match:
            return 0, 0
        col_letters, row_text = match.groups()
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - 64)
        return int(row_text), col

    @staticmethod
    def _shared_strings(package: ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in package.namelist():
            return []
        root = ET.fromstring(package.read("xl/sharedStrings.xml"))
        strings = []
        for si in root.findall(f"{{{MAIN_NS}}}si"):
            strings.append("".join(node.text or "" for node in si.iter(f"{{{MAIN_NS}}}t")))
        return strings

    @classmethod
    def _count_activity_rows(cls, root: ET.Element, shared_strings: list[str]) -> int:
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            return 0
        activity_col = None
        header_row = None
        for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
            row_num = int(row.attrib.get("r", "0"))
            if row_num > 30:
                break
            for cell in row.findall(f"{{{MAIN_NS}}}c"):
                value = cls._cell_text(cell, shared_strings).strip().lower()
                if value in cls.ACTIVITY_ID_HEADERS:
                    activity_col = re.match(r"[A-Z]+", cell.attrib.get("r", ""))[0]
                    header_row = row_num
                    break
            if activity_col:
                break
        if not activity_col or header_row is None:
            return 0

        activity_ids: set[str] = set()
        for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
            row_num = int(row.attrib.get("r", "0"))
            if row_num <= header_row:
                continue
            for cell in row.findall(f"{{{MAIN_NS}}}c"):
                ref = cell.attrib.get("r", "")
                match = re.match(r"[A-Z]+", ref)
                if match and match[0] == activity_col:
                    text = cls._cell_text(cell, shared_strings).strip()
                    if text:
                        activity_ids.add(text)
                    break
        return len(activity_ids)

    @staticmethod
    def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
        cell_type = cell.attrib.get("t")
        if cell_type == "inlineStr":
            inline = cell.find(f"{{{MAIN_NS}}}is")
            if inline is None:
                return ""
            return "".join(node.text or "" for node in inline.iter(f"{{{MAIN_NS}}}t"))
        value = cell.find(f"{{{MAIN_NS}}}v")
        if value is None or value.text is None:
            return ""
        if cell_type == "s":
            try:
                return shared_strings[int(value.text)]
            except (ValueError, IndexError):
                return ""
        return value.text
