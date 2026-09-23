from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]

PLANNER_HEADERS = ("Ψυχ_Ώρα", "Ψυχ_Ημέρες", "Ψυχ_Ψυχολόγος")
SETTINGS_HEADER = "PSYCHOLOGISTS"
VALIDATION_LAST_ROW = 500


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


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


def _copy_header_style_and_width(ws, source_col: int, target_col: int) -> None:
    # xlPasteFormats = -4122. Copy only the header cell so we do not inflate
    # the used range or touch data/validation in any existing source column.
    ws.Cells(1, source_col).Copy()
    ws.Cells(1, target_col).PasteSpecial(Paste=-4122)
    ws.Columns(target_col).ColumnWidth = ws.Columns(source_col).ColumnWidth


def _replace_workbook_name(workbook, name: str, refers_to: str) -> None:
    try:
        workbook.Names(name).Delete()
    except Exception:
        pass
    workbook.Names.Add(Name=name, RefersTo=refers_to)


def _set_list_validation(ws, column: int, formula1: str) -> None:
    target = ws.Range(ws.Cells(2, column), ws.Cells(VALIDATION_LAST_ROW, column))
    try:
        target.Validation.Delete()
    except Exception:
        pass
    # xlValidateList = 3, xlValidAlertStop = 1
    target.Validation.Add(Type=3, AlertStyle=1, Formula1=formula1)
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.ShowError = True
    target.Validation.ErrorTitle = "Μη έγκυρη επιλογή"
    target.Validation.ErrorMessage = "Επίλεξε τιμή από την αναπτυσσόμενη λίστα."


def _apply_dropdowns(workbook, planner_cols: tuple[int, int, int], settings_col: int) -> None:
    planner = workbook.Worksheets("PATIENT_PLANNER")

    # Dynamic workbook names keep the dropdowns in sync with SETTINGS as rows
    # are added later. Header row is excluded; when a list is empty, row 2 is
    # still a valid blank target so Excel keeps the dropdown definition alive.
    _replace_workbook_name(
        workbook,
        "PSYCH_HOURS_LIST",
        "=SETTINGS!$B$2:INDEX(SETTINGS!$B:$B,MAX(2,COUNTA(SETTINGS!$B:$B)))",
    )
    _replace_workbook_name(
        workbook,
        "PSYCH_DAYS_LIST",
        "=SETTINGS!$D$2:INDEX(SETTINGS!$D:$D,MAX(2,COUNTA(SETTINGS!$D:$D)))",
    )
    settings_letter = workbook.Application.ConvertFormula(
        workbook.Worksheets("SETTINGS").Cells(1, settings_col).Address,
        1,
        1,
        1,
    ).replace("$1", "")
    _replace_workbook_name(
        workbook,
        "PSYCHOLOGISTS_LIST",
        (
            f"=SETTINGS!{settings_letter}$2:INDEX(SETTINGS!{settings_letter}:{settings_letter},"
            f"MAX(2,COUNTA(SETTINGS!{settings_letter}:{settings_letter})))"
        ),
    )

    _set_list_validation(planner, planner_cols[0], "=PSYCH_HOURS_LIST")
    _set_list_validation(planner, planner_cols[1], "=PSYCH_DAYS_LIST")
    _set_list_validation(planner, planner_cols[2], "=PSYCHOLOGISTS_LIST")


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

    _apply_dropdowns(workbook, planner_cols, settings_col)
    workbook.Application.CutCopyMode = False
    return planner_cols, settings_col


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
    source_vba = _has_vba(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    excel = None
    workbook = None
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
        if output.exists():
            try:
                output.unlink()
            except OSError:
                pass
        print(f"SAFETY STOP: {exc}")
        return 2
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

    source_after = _sha256(source)
    source_unchanged = source_before == source_after
    vba_preserved = source_vba == _has_vba(output) and _has_vba(output)

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
    print("Dropdowns added: Ψυχ_Ώρα, Ψυχ_Ημέρες, Ψυχ_Ψυχολόγος")
    print(f"Source unchanged: {source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print("NOTE: no final visual timetable/layout changes are made by this tool.")
    return 0 if source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
