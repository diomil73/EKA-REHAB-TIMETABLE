from __future__ import annotations

import argparse
from pathlib import Path
import sys
from zipfile import ZipFile


REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel import (  # noqa: E402
    CellPatch,
    WriteIntent,
    apply_write_plan_to_copy,
    build_write_plan,
)

SOURCE_NAME = "Rehab_Center_System_v27_1.xlsm"
OUTPUT_NAME = "Rehab_Center_System_v27_1_NATIVE_SMOKE.xlsm"
TARGET_SHEET = "THERAPIST_DAILY"
TARGET_CELL = "B8"
MARKER = "PYTHON SMOKE TEST"


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


def read_excel_cell(path: Path, sheet: str, cell: str):
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise SystemExit("pywin32 is not installed. Run: python -m pip install pywin32") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(
            str(path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
        )
        return workbook.Worksheets(sheet).Range(cell).Value
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create the first safe native-Excel write-back preview copy."
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

    source = args.source.resolve() if args.source else find_source(REPO_ROOT)
    if not source.exists():
        raise SystemExit(f"SOURCE NOT FOUND: {source}")
    if source.suffix.lower() != ".xlsm":
        raise SystemExit("SOURCE MUST BE .xlsm")

    output = (
        args.output.resolve()
        if args.output
        else (REPO_ROOT / "Excel" / "previews" / OUTPUT_NAME).resolve()
    )

    if output.exists() and not args.overwrite:
        raise SystemExit(
            f"OUTPUT ALREADY EXISTS: {output}\n"
            "Delete it or rerun with --overwrite."
        )

    before_value = read_excel_cell(source, TARGET_SHEET, TARGET_CELL)
    if before_value not in (None, ""):
        raise SystemExit(
            f"SAFETY STOP: expected {TARGET_SHEET}!{TARGET_CELL} to be empty in v27.1, "
            f"but found {before_value!r}. No file was written."
        )

    if not has_vba_project(source):
        raise SystemExit("SAFETY STOP: source workbook does not contain xl/vbaProject.bin")

    plan = build_write_plan(
        source,
        [
            CellPatch(
                sheet=TARGET_SHEET,
                cell=TARGET_CELL,
                value=MARKER,
                intent=WriteIntent.PRESENTATION,
                wrap_text=True,
                min_font_size=10,
                source_tag="native_smoke_test",
            )
        ],
    )

    report = apply_write_plan_to_copy(
        plan,
        output,
        overwrite=args.overwrite,
    )

    after_value = read_excel_cell(output, TARGET_SHEET, TARGET_CELL)
    if after_value != MARKER:
        output.unlink(missing_ok=True)
        raise SystemExit(
            f"FAILED: preview cell read-back was {after_value!r}, expected {MARKER!r}. "
            "Preview copy was deleted."
        )

    if not has_vba_project(output):
        output.unlink(missing_ok=True)
        raise SystemExit(
            "FAILED: VBA project was not preserved in the preview. Preview copy was deleted."
        )

    print("SMOKE OK")
    print(f"Source: {source}")
    print(f"Preview: {output}")
    print(f"Write: {TARGET_SHEET}!{TARGET_CELL} = {MARKER}")
    print(f"Source unchanged: {report.source_unchanged}")
    print("VBA preserved: True")
    print("NEXT: open ONLY the preview workbook in Excel and inspect it manually.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
