from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile


REPO_ROOT = Path(__file__).resolve().parents[2]
SOURCE_NAME = "Rehab_Center_System_v27_1.xlsm"
OUTPUT_NAME = "Rehab_Center_System_v27_1_NATIVE_SMOKE.xlsm"
TARGET_SHEET = "THERAPIST_DAILY"
MARKER = "PYTHON SMOKE TEST"

# Only the confirmed patient-grid cells of THERAPIST_DAILY are candidates.
# Search from the lower/right side first so the smoke marker is less likely to
# cover an operationally important cell in the preview copy.
GRID_BLOCKS = (
    (12, 18, 2, 10),  # B12:J18
    (2, 8, 2, 10),    # B2:J8
)


def _column_letters(number: int) -> str:
    chars: list[str] = []
    while number:
        number, rem = divmod(number - 1, 26)
        chars.append(chr(ord("A") + rem))
    return "".join(reversed(chars))


def candidate_cells() -> tuple[str, ...]:
    cells: list[str] = []
    for first_row, last_row, first_col, last_col in GRID_BLOCKS:
        for row in range(last_row, first_row - 1, -1):
            for col in range(last_col, first_col - 1, -1):
                cells.append(f"{_column_letters(col)}{row}")
    return tuple(cells)


def file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def find_source(root: Path) -> Path:
    candidates = [
        p.resolve()
        for p in root.rglob(SOURCE_NAME)
        if "previews" not in {part.lower() for part in p.parts}
    ]
    if not candidates:
        raise SystemExit(
            f"SOURCE NOT FOUND: {SOURCE_NAME}\n"
            "Put the v27.1 baseline workbook somewhere inside this repository."
        )

    baseline_candidates = [
        p for p in candidates if "baseline" in {part.lower() for part in p.parts}
    ]
    if len(baseline_candidates) == 1:
        return baseline_candidates[0]
    if len(candidates) == 1:
        return candidates[0]

    options = "\n".join(f"  - {p}" for p in candidates)
    raise SystemExit(
        "MULTIPLE SOURCES FOUND. Run again with --source <path>.\n" + options
    )


def has_vba_project(path: Path) -> bool:
    with ZipFile(path) as archive:
        names = {name.lower() for name in archive.namelist()}
    return "xl/vbaproject.bin" in names


def _get_excel_module():
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise SystemExit(
            "pywin32 is not installed. Run: python -m pip install pywin32"
        ) from exc
    return win32com.client


def find_safe_empty_cell(path: Path) -> str:
    """Find one truly empty, non-formula, non-merged cell in confirmed grids."""

    win32 = _get_excel_module()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(
            str(path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        worksheet = workbook.Worksheets(TARGET_SHEET)

        for address in candidate_cells():
            cell = worksheet.Range(address)
            value = cell.Value
            try:
                has_formula = bool(cell.HasFormula)
            except Exception:
                has_formula = False
            try:
                is_merged = bool(cell.MergeCells)
            except Exception:
                is_merged = False

            if value in (None, "") and not has_formula and not is_merged:
                return address
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()

    raise SystemExit(
        "SAFETY STOP: no empty non-formula cell was found in the confirmed "
        "THERAPIST_DAILY grids B2:J8 or B12:J18. No file was written."
    )


def write_marker_to_copy(source: Path, output: Path, cell_address: str) -> None:
    win32 = _get_excel_module()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(
            str(output.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )
        cell = workbook.Worksheets(TARGET_SHEET).Range(cell_address)
        cell.Value = MARKER
        cell.WrapText = True
        try:
            if float(cell.Font.Size) < 10:
                cell.Font.Size = 10
        except (TypeError, ValueError):
            cell.Font.Size = 10
        workbook.Save()
    except Exception:
        output.unlink(missing_ok=True)
        raise
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


def read_excel_cell(path: Path, sheet: str, cell: str):
    win32 = _get_excel_module()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(
            str(path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        return workbook.Worksheets(sheet).Range(cell).Value
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create a safe native-Excel smoke preview using an automatically "
            "selected empty THERAPIST_DAILY cell."
        )
    )
    parser.add_argument("--source", type=Path, help="Optional explicit source .xlsm path")
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional output .xlsm path (default: Excel/previews/native smoke copy)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace an existing smoke preview copy",
    )
    args = parser.parse_args()

    if sys.platform != "win32":
        raise SystemExit("NOT READY: this smoke test requires Windows + Microsoft Excel")

    source = args.source.resolve() if args.source else find_source(REPO_ROOT)
    if not source.exists():
        raise SystemExit(f"SOURCE NOT FOUND: {source}")
    if source.suffix.lower() != ".xlsm":
        raise SystemExit("SOURCE MUST BE .xlsm")
    if not has_vba_project(source):
        raise SystemExit("SAFETY STOP: source workbook does not contain xl/vbaProject.bin")

    output = (
        args.output.resolve()
        if args.output
        else (REPO_ROOT / "Excel" / "previews" / OUTPUT_NAME).resolve()
    )
    if source == output:
        raise SystemExit("SAFETY STOP: output path must be different from source")
    if output.exists() and not args.overwrite:
        raise SystemExit(
            f"OUTPUT ALREADY EXISTS: {output}\n"
            "Delete it or rerun with --overwrite."
        )

    source_hash_before = file_sha256(source)
    target_cell = find_safe_empty_cell(source)

    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    try:
        write_marker_to_copy(source, output, target_cell)

        after_value = read_excel_cell(output, TARGET_SHEET, target_cell)
        if after_value != MARKER:
            raise RuntimeError(
                f"preview cell read-back was {after_value!r}, expected {MARKER!r}"
            )
        if not has_vba_project(output):
            raise RuntimeError("VBA project was not preserved in the preview")

        source_hash_after = file_sha256(source)
        if source_hash_before != source_hash_after:
            raise RuntimeError("source workbook changed during the smoke test")
    except Exception as exc:
        output.unlink(missing_ok=True)
        raise SystemExit(f"FAILED: {exc}. Preview copy was deleted.") from exc

    print("SMOKE OK")
    print(f"Source: {source}")
    print(f"Preview: {output}")
    print(f"Auto-selected write: {TARGET_SHEET}!{target_cell} = {MARKER}")
    print("Source unchanged: True")
    print("VBA preserved: True")
    print("NEXT: open ONLY the preview workbook in Excel and inspect it manually.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
