from __future__ import annotations

import argparse
from hashlib import sha256
import os
from pathlib import Path
import posixpath
import re
import shutil
import sys
import tempfile
from uuid import uuid4
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

REPO_ROOT = Path(__file__).resolve().parents[2]

PLANNER_HEADERS = ("Ψυχ_Ώρα", "Ψυχ_Ημέρες", "Ψυχ_Ψυχολόγος")
SETTINGS_HEADER = "PSYCHOLOGISTS"
VALIDATION_LAST_ROW = 500

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _zip_member_sha256(path: Path, member: str) -> str | None:
    with ZipFile(path) as archive:
        try:
            data = archive.read(member)
        except KeyError:
            return None
    return sha256(data).hexdigest()


def _has_vba(path: Path) -> bool:
    return _zip_member_sha256(path, "xl/vbaProject.bin") is not None


def _last_header_column(ws) -> int:
    # xlToLeft = -4159
    return int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)


def _header_map(ws) -> dict[str, int]:
    last_col = _last_header_column(ws)
    result: dict[str, int] = {}
    for col in range(1, last_col + 1):
        value = ws.Cells(1, col).Value
        if value is None:
            continue
        text = str(value).strip()
        if text:
            result[text] = col
    return result


def _column_letter(column: int) -> str:
    if column < 1:
        raise ValueError("Excel column index must be >= 1")
    letters: list[str] = []
    value = column
    while value:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def _copy_header_style_and_width(ws, source_col: int, target_col: int) -> None:
    # xlPasteFormats = -4122. Copy only the header cell so we do not inflate
    # the used range or touch data/validation in any existing source column.
    ws.Cells(1, source_col).Copy()
    ws.Cells(1, target_col).PasteSpecial(Paste=-4122)
    ws.Columns(target_col).ColumnWidth = ws.Columns(source_col).ColumnWidth


def _apply_schema(workbook) -> tuple[tuple[int, int, int], int]:
    planner = workbook.Worksheets("PATIENT_PLANNER")
    settings = workbook.Worksheets("SETTINGS")

    planner_headers = _header_map(planner)
    existing = [header in planner_headers for header in PLANNER_HEADERS]
    if any(existing) and not all(existing):
        raise RuntimeError(
            "PATIENT_PLANNER contains a partial psychology schema; refusing to guess."
        )

    if all(existing):
        planner_cols = tuple(planner_headers[header] for header in PLANNER_HEADERS)
    else:
        start = _last_header_column(planner) + 1
        planner_cols = (start, start + 1, start + 2)

        # No existing source column is moved. We only borrow header styling and
        # widths so the preview remains legible; final UI/layout is out of scope.
        if "ΕΦΑ_Ώρα" in planner_headers:
            _copy_header_style_and_width(
                planner, planner_headers["ΕΦΑ_Ώρα"], planner_cols[0]
            )
        if "ΕΦΑ_Ημέρες" in planner_headers:
            _copy_header_style_and_width(
                planner, planner_headers["ΕΦΑ_Ημέρες"], planner_cols[1]
            )
        if "ΦΘ_Θεραπευτής" in planner_headers:
            _copy_header_style_and_width(
                planner, planner_headers["ΦΘ_Θεραπευτής"], planner_cols[2]
            )

        for col, header in zip(planner_cols, PLANNER_HEADERS):
            planner.Cells(1, col).Value = header

    settings_headers = _header_map(settings)
    if SETTINGS_HEADER in settings_headers:
        settings_col = settings_headers[SETTINGS_HEADER]
    else:
        settings_col = _last_header_column(settings) + 1
        if "THERAPISTS_FTH" in settings_headers:
            _copy_header_style_and_width(
                settings, settings_headers["THERAPISTS_FTH"], settings_col
            )
        settings.Cells(1, settings_col).Value = SETTINGS_HEADER

    workbook.Application.CutCopyMode = False
    return planner_cols, settings_col


def _sheet_xml_member(path: Path, sheet_name: str) -> str:
    ns = {"m": MAIN_NS, "r": DOC_REL_NS, "pr": PKG_REL_NS}
    with ZipFile(path) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    rel_id = None
    for sheet in workbook_root.findall("m:sheets/m:sheet", ns):
        if sheet.attrib.get("name") == sheet_name:
            rel_id = sheet.attrib.get(f"{{{DOC_REL_NS}}}id")
            break
    if not rel_id:
        raise RuntimeError(f"Could not resolve sheet relationship for {sheet_name}")

    target = None
    for rel in rels_root.findall("pr:Relationship", ns):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib.get("Target")
            break
    if not target:
        raise RuntimeError(f"Could not resolve sheet XML target for {sheet_name}")

    if target.startswith("/"):
        member = target.lstrip("/")
    else:
        member = posixpath.normpath(posixpath.join("xl", target))
    return member


def _replace_zip_member(path: Path, member: str, data: bytes) -> None:
    fd, tmp_name = tempfile.mkstemp(
        prefix=path.stem + "_", suffix=".tmp", dir=str(path.parent)
    )
    os.close(fd)
    tmp_path = Path(tmp_name)

    try:
        with ZipFile(path, "r") as source_zip, ZipFile(
            tmp_path, "w", compression=ZIP_DEFLATED
        ) as target_zip:
            for info in source_zip.infolist():
                payload = data if info.filename == member else source_zip.read(info.filename)
                target_zip.writestr(info, payload)
        os.replace(tmp_path, path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def _new_uid_text(old: str) -> str:
    value = str(uuid4()).upper()
    if old.startswith("{") and old.endswith("}"):
        return "{" + value + "}"
    return value


def _clone_validation_node(
    node: str,
    *,
    target_column: int,
    formula_range_replacement: tuple[str, str] | None = None,
) -> str:
    target_letter = _column_letter(target_column)

    def uid_repl(match: re.Match[str]) -> str:
        return match.group(1) + _new_uid_text(match.group(2)) + match.group(3)

    clone = re.sub(
        r'(\b(?:[A-Za-z_][\w.-]*:)?uid=")(\{?[0-9A-Fa-f-]{36}\}?)(\")',
        uid_repl,
        node,
    )

    sqref_pattern = re.compile(
        r'(<[A-Za-z_][\w.-]*:sqref>).*?(</[A-Za-z_][\w.-]*:sqref>)',
        re.DOTALL,
    )
    if not sqref_pattern.search(clone):
        raise RuntimeError("Could not find sqref inside existing extended validation")
    clone = sqref_pattern.sub(
        lambda match: (
            f"{match.group(1)}{target_letter}2:"
            f"{target_letter}{VALIDATION_LAST_ROW}{match.group(2)}"
        ),
        clone,
        count=1,
    )

    if formula_range_replacement is not None:
        old, new = formula_range_replacement
        if old not in clone:
            raise RuntimeError(
                f"Could not find expected provider range {old} in validation XML"
            )
        clone = clone.replace(old, new, 1)

    return clone


def _patch_extended_validations(
    path: Path,
    planner_cols: tuple[int, int, int],
    settings_col: int,
) -> None:
    """Clone the workbook's existing x14 data validations.

    This workbook stores cross-sheet list validation in Excel's extended x14
    validation XML. Excel COM can read those existing validations, but rejects
    creating an equivalent cross-sheet validation through Validation.Add.
    Cloning the existing validated XML keeps the workbook's own mechanism.
    """

    member = _sheet_xml_member(path, "PATIENT_PLANNER")
    with ZipFile(path) as archive:
        raw = archive.read(member)

    try:
        xml = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise RuntimeError("PATIENT_PLANNER XML is not UTF-8") from exc

    block_pattern = re.compile(
        r'(?P<open><(?P<prefix>[A-Za-z_][\w.-]*):dataValidations\b[^>]*>)'
        r'(?P<body>.*?)'
        r'(?P<close></(?P=prefix):dataValidations>)',
        re.DOTALL,
    )

    chosen = None
    for match in block_pattern.finditer(xml):
        body = match.group("body")
        if (
            "$B$2:$B$20" in body
            and "$D$2:$D$100" in body
            and "$A$2:$A$100" in body
        ):
            chosen = match
            break
    if chosen is None:
        raise RuntimeError(
            "Could not find the workbook's existing extended PATIENT_PLANNER "
            "validation block to clone."
        )

    prefix = chosen.group("prefix")
    node_pattern = re.compile(
        rf'<{re.escape(prefix)}:dataValidation\b.*?'
        rf'</{re.escape(prefix)}:dataValidation>',
        re.DOTALL,
    )
    nodes = node_pattern.findall(chosen.group("body"))

    def find_source(range_text: str) -> str:
        for node in nodes:
            if "SETTINGS!" in node and range_text in node:
                return node
        raise RuntimeError(
            f"Could not find existing validation source for SETTINGS!{range_text}"
        )

    hours_source = find_source("$B$2:$B$20")
    days_source = find_source("$D$2:$D$100")
    provider_source = find_source("$A$2:$A$100")

    settings_letter = _column_letter(settings_col)
    clones = [
        _clone_validation_node(
            hours_source,
            target_column=planner_cols[0],
        ),
        _clone_validation_node(
            days_source,
            target_column=planner_cols[1],
        ),
        _clone_validation_node(
            provider_source,
            target_column=planner_cols[2],
            formula_range_replacement=(
                "$A$2:$A$100",
                f"${settings_letter}$2:${settings_letter}$100",
            ),
        ),
    ]

    new_body = chosen.group("body").rstrip() + "\n" + "\n".join(clones) + "\n"
    new_count = len(node_pattern.findall(new_body))

    open_tag = chosen.group("open")
    if re.search(r'\bcount="\d+"', open_tag):
        open_tag = re.sub(
            r'\bcount="\d+"', f'count="{new_count}"', open_tag, count=1
        )
    else:
        open_tag = open_tag[:-1] + f' count="{new_count}">'

    replacement = open_tag + new_body + chosen.group("close")
    patched = (
        xml[: chosen.start()] + replacement + xml[chosen.end() :]
    ).encode("utf-8")

    _replace_zip_member(path, member, patched)


def _normalize_formula(value: object) -> str:
    return str(value or "").replace("'", "").replace(" ", "").upper()


def _verify_dropdowns(
    path: Path,
    planner_cols: tuple[int, int, int],
    settings_col: int,
) -> None:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        raise RuntimeError("pywin32 is required for dropdown verification")

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        planner = workbook.Worksheets("PATIENT_PLANNER")

        settings_letter = _column_letter(settings_col)
        expected = (
            "=SETTINGS!$B$2:$B$20",
            "=SETTINGS!$D$2:$D$100",
            f"=SETTINGS!${settings_letter}$2:${settings_letter}$100",
        )

        for header, col, expected_formula in zip(
            PLANNER_HEADERS, planner_cols, expected
        ):
            cell = planner.Cells(2, col)
            validation_type = int(cell.Validation.Type)
            formula1 = str(cell.Validation.Formula1)
            if validation_type != 3:
                raise RuntimeError(
                    f"{header} validation type is {validation_type}, expected list type 3"
                )
            if _normalize_formula(formula1) != _normalize_formula(expected_formula):
                raise RuntimeError(
                    f"{header} Formula1 is {formula1!r}, expected {expected_formula!r}"
                )
            if not bool(cell.Validation.InCellDropdown):
                raise RuntimeError(f"{header} dropdown is not enabled")
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def _close_excel(workbook, excel) -> None:
    if workbook is not None:
        try:
            workbook.Close(SaveChanges=False)
        except Exception:
            pass
    if excel is not None:
        try:
            excel.Quit()
        except Exception:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create a safe .xlsm preview that appends psychology source columns, "
            "their Excel dropdown validations, and a PSYCHOLOGISTS registry. "
            "The source workbook is never edited."
        )
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "PSYCHOLOGY_SCHEMA_PREVIEW.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()

    if sys.platform != "win32":
        print("SAFETY STOP: this preview requires Windows with Microsoft Excel.")
        return 2
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2
    if source == output:
        print("SAFETY STOP: output must be different from source workbook.")
        return 2
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        print("SAFETY STOP: source and output must both be .xlsm files.")
        return 2
    if output.exists() and not args.overwrite:
        print(f"SAFETY STOP: output already exists: {output}")
        return 2

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 is required (pip install pywin32).")
        return 2

    source_before = _sha256(source)
    source_vba_hash = _zip_member_sha256(source, "xl/vbaProject.bin")
    output.parent.mkdir(parents=True, exist_ok=True)

    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    planner_cols: tuple[int, int, int] | None = None
    settings_col: int | None = None
    excel = None
    workbook = None
    schema_error: Exception | None = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)

        sheet_names = {
            str(workbook.Worksheets(i).Name)
            for i in range(1, workbook.Worksheets.Count + 1)
        }
        required = {"PATIENT_PLANNER", "SETTINGS"}
        missing = sorted(required - sheet_names)
        if missing:
            raise RuntimeError("Missing required sheet(s): " + ", ".join(missing))

        planner_cols, settings_col = _apply_schema(workbook)
        workbook.Save()
    except Exception as exc:
        schema_error = exc
    finally:
        _close_excel(workbook, excel)

    if schema_error is not None:
        if output.exists():
            output.unlink()
        print(f"SAFETY STOP: schema stage failed: {schema_error}")
        return 2

    assert planner_cols is not None
    assert settings_col is not None

    try:
        _patch_extended_validations(output, planner_cols, settings_col)
    except Exception as exc:
        if output.exists():
            output.unlink()
        print(f"SAFETY STOP: validation XML stage failed: {exc}")
        return 2

    try:
        _verify_dropdowns(output, planner_cols, settings_col)
    except Exception as exc:
        if output.exists():
            output.unlink()
        print(f"SAFETY STOP: validation verification failed: {exc}")
        return 2

    source_after = _sha256(source)
    source_unchanged = source_before == source_after
    output_vba_hash = _zip_member_sha256(output, "xl/vbaProject.bin")
    vba_preserved = (
        source_vba_hash is not None
        and output_vba_hash is not None
        and source_vba_hash == output_vba_hash
    )

    print("PSYCHOLOGY SCHEMA PREVIEW OK")
    print(f"Source: {source}")
    print(f"Preview: {output}")
    print(
        "PATIENT_PLANNER appended columns: "
        + ", ".join(
            f"{header}=col {col}" for header, col in zip(PLANNER_HEADERS, planner_cols)
        )
    )
    print(f"SETTINGS appended registry: {SETTINGS_HEADER}=col {settings_col}")
    print("Dropdowns verified: Ψυχ_Ώρα, Ψυχ_Ημέρες, Ψυχ_Ψυχολόγος")
    print(f"Source unchanged: {source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print("NOTE: no final visual timetable/layout changes are made by this tool.")
    return 0 if source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
