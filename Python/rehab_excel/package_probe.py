from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path, PurePosixPath
import re
import posixpath
import xml.etree.ElementTree as ET
from zipfile import ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


@dataclass(frozen=True)
class SheetPackageInfo:
    name: str
    part: str
    dimension: str | None


class WorkbookPackageProbe:
    """Read workbook package metadata without opening or saving in Excel.

    The probe reads the .xlsx/.xlsm ZIP package directly. It is intentionally
    read-only and is used only to validate the known workbook layout before a
    future write-back is allowed.
    """

    def __init__(self, workbook_path: str | Path):
        self.path = Path(workbook_path)
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        self._sheet_parts: dict[str, str] | None = None
        self._shared_strings: tuple[str, ...] | None = None

    @staticmethod
    def _join_part(base_part: str, target: str) -> str:
        base = str(PurePosixPath(base_part).parent)
        return posixpath.normpath(posixpath.join(base, target))

    def _load_sheet_parts(self) -> dict[str, str]:
        if self._sheet_parts is not None:
            return self._sheet_parts

        with ZipFile(self.path) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rel_map = {
                rel.attrib["Id"]: self._join_part("xl/workbook.xml", rel.attrib["Target"])
                for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
            }
            result: dict[str, str] = {}
            sheets = workbook.find(f"{{{MAIN_NS}}}sheets")
            if sheets is not None:
                for sheet in sheets:
                    rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
                    result[sheet.attrib["name"]] = rel_map[rel_id]

        self._sheet_parts = result
        return result

    @property
    def sheet_names(self) -> tuple[str, ...]:
        return tuple(self._load_sheet_parts().keys())

    def sheet_part(self, sheet_name: str) -> str:
        try:
            return self._load_sheet_parts()[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Unknown workbook sheet: {sheet_name}") from exc

    def _load_shared_strings(self) -> tuple[str, ...]:
        if self._shared_strings is not None:
            return self._shared_strings

        values: list[str] = []
        with ZipFile(self.path) as archive:
            if "xl/sharedStrings.xml" in archive.namelist():
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                for item in root.findall(f"{{{MAIN_NS}}}si"):
                    values.append(
                        "".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t"))
                    )
        self._shared_strings = tuple(values)
        return self._shared_strings

    def sheet_info(self, sheet_name: str) -> SheetPackageInfo:
        part = self.sheet_part(sheet_name)
        with ZipFile(self.path) as archive:
            root = ET.fromstring(archive.read(part))
        dimension = root.find(f"{{{MAIN_NS}}}dimension")
        return SheetPackageInfo(
            name=sheet_name,
            part=part,
            dimension=dimension.attrib.get("ref") if dimension is not None else None,
        )

    def cell_values(self, sheet_name: str) -> dict[str, object]:
        """Return cached/display cell values keyed by A1 address.

        Formula text is not evaluated. Existing cached values are returned,
        which is sufficient for the layout/header mapping performed here.
        """

        shared = self._load_shared_strings()
        part = self.sheet_part(sheet_name)
        with ZipFile(self.path) as archive:
            root = ET.fromstring(archive.read(part))

        values: dict[str, object] = {}
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            return values

        for row in sheet_data:
            for cell in row.findall(f"{{{MAIN_NS}}}c"):
                address = cell.attrib.get("r")
                if not address:
                    continue
                kind = cell.attrib.get("t")
                if kind == "inlineStr":
                    value: object = "".join(
                        node.text or "" for node in cell.iter(f"{{{MAIN_NS}}}t")
                    )
                else:
                    value_node = cell.find(f"{{{MAIN_NS}}}v")
                    if value_node is None:
                        continue
                    raw = value_node.text or ""
                    if kind == "s":
                        value = shared[int(raw)]
                    elif kind == "b":
                        value = raw == "1"
                    else:
                        value = raw
                values[address.upper()] = value
        return values

    def _sheet_relationships(self, sheet_name: str) -> dict[str, tuple[str, str]]:
        part = PurePosixPath(self.sheet_part(sheet_name))
        rels_part = str(part.parent / "_rels" / f"{part.name}.rels")
        with ZipFile(self.path) as archive:
            if rels_part not in archive.namelist():
                return {}
            root = ET.fromstring(archive.read(rels_part))
        return {
            rel.attrib["Id"]: (
                rel.attrib.get("Type", ""),
                self._join_part(str(part), rel.attrib["Target"]),
            )
            for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship")
        }

    def drawing_macros(self, sheet_name: str) -> tuple[str, ...]:
        """Return macro names attached to DrawingML shapes on a worksheet."""

        relationships = self._sheet_relationships(sheet_name)
        if not relationships:
            return ()

        part = self.sheet_part(sheet_name)
        with ZipFile(self.path) as archive:
            sheet_root = ET.fromstring(archive.read(part))
            drawing = sheet_root.find(f"{{{MAIN_NS}}}drawing")
            if drawing is None:
                return ()
            rel_id = drawing.attrib.get(f"{{{REL_NS}}}id")
            if not rel_id or rel_id not in relationships:
                return ()
            drawing_part = relationships[rel_id][1]
            root = ET.fromstring(archive.read(drawing_part))

        macros: list[str] = []
        for shape in root.iter(f"{{{DRAW_NS}}}sp"):
            macro = (shape.attrib.get("macro") or "").strip()
            if macro:
                macros.append(re.sub(r"^\[\d+\]!", "", macro))
        return tuple(macros)

    def legacy_button_macros(self, sheet_name: str) -> tuple[str, ...]:
        """Return VBA macros attached to legacy VML form-control buttons."""

        relationships = self._sheet_relationships(sheet_name)
        if not relationships:
            return ()

        part = self.sheet_part(sheet_name)
        with ZipFile(self.path) as archive:
            sheet_root = ET.fromstring(archive.read(part))
            legacy = sheet_root.find(f"{{{MAIN_NS}}}legacyDrawing")
            if legacy is None:
                return ()
            rel_id = legacy.attrib.get(f"{{{REL_NS}}}id")
            if not rel_id or rel_id not in relationships:
                return ()
            vml_part = relationships[rel_id][1]
            text = archive.read(vml_part).decode("utf-8", errors="ignore")

        macros = re.findall(r"<x:FmlaMacro>\s*(.*?)\s*</x:FmlaMacro>", text)
        return tuple(re.sub(r"^\[\d+\]!", "", macro) for macro in macros)
