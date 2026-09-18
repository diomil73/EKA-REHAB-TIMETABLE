from __future__ import annotations

import argparse
import sys
from datetime import time
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.layout_contract import resolve_therapist_daily_cell  # noqa: E402
from rehab_excel.native_excel import apply_write_plan_to_copy  # noqa: E402
from rehab_excel.writeback import CellPatch, TextRunPatch, build_write_plan  # noqa: E402

PATIENT_NAME = "ΖΑΛΟΚΩΣΤΑΣ"
ORIGINAL_PROVIDER = "Αργέντος"
REPLACEMENT_PROVIDER = "Φιλιππούσης"
SLOT_TIME = time(12, 15)


def _clean(value: object | None) -> str:
    if value is None:
        return ""
    return str(value).strip()


def append_unique_line(existing: str, line: str) -> str:
    """Append a display line unless it already appears as a full line."""

    current_lines = [item.strip() for item in existing.splitlines() if item.strip()]
    if line in current_lines:
        return "\n".join(current_lines)
    current_lines.append(line)
    return "\n".join(current_lines)


def line_run_for_text(
    full_text: str,
    needle: str,
    *,
    strike: bool,
    italic: bool = False,
    font_role: str | None,
):
    """Return a 1-based rich-text run covering the line containing needle."""

    cursor = 0
    for line in full_text.splitlines(keepends=True):
        visible = line.rstrip("\r\n")
        if needle in visible:
            return TextRunPatch(
                start=cursor + 1,
                length=len(visible),
                strike_through=strike,
                italic=italic,
                font_role=font_role,
            )
        cursor += len(line)
    raise ValueError(f"{needle!r} not found in cell text")


def _read_cell_text(workbook_path: Path, cell_address: str) -> str:
    wb = load_workbook(
        workbook_path,
        read_only=True,
        data_only=False,
        keep_vba=True,
    )
    try:
        return _clean(wb["THERAPIST_DAILY"][cell_address].value)
    finally:
        wb.close()


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def build_preview_plan(source: Path):
    original_cell = resolve_therapist_daily_cell(
        source, ORIGINAL_PROVIDER, SLOT_TIME
    )
    replacement_cell = resolve_therapist_daily_cell(
        source, REPLACEMENT_PROVIDER, SLOT_TIME
    )

    original_before = _read_cell_text(source, original_cell)
    replacement_before = _read_cell_text(source, replacement_cell)

    original_after = append_unique_line(original_before, PATIENT_NAME)
    replacement_note = (
        f"→ {REPLACEMENT_PROVIDER} {SLOT_TIME.strftime('%H:%M')}"
    )
    original_after = append_unique_line(original_after, replacement_note)
    replacement_after = append_unique_line(replacement_before, PATIENT_NAME)

    original_run = line_run_for_text(
        original_after,
        PATIENT_NAME,
        strike=False,
        italic=True,
        font_role="muted",
    )
    replacement_note_run = line_run_for_text(
        original_after,
        replacement_note,
        strike=False,
        font_role="default",
    )
    replacement_run = line_run_for_text(
        replacement_after,
        PATIENT_NAME,
        strike=False,
        font_role="default",
    )

    patches = (
        CellPatch(
            sheet="THERAPIST_DAILY",
            cell=original_cell,
            value=original_after,
            wrap_text=True,
            min_font_size=10,
            fill_role="infectious_yellow",
            border_role="infectious_yellow",
            text_runs=(original_run, replacement_note_run),
            source_tag="operational_preview:original",
        ),
        CellPatch(
            sheet="THERAPIST_DAILY",
            cell=replacement_cell,
            value=replacement_after,
            wrap_text=True,
            min_font_size=10,
            fill_role="infectious_yellow",
            border_role="infectious_yellow",
            text_runs=(replacement_run,),
            source_tag="operational_preview:replacement",
        ),
    )

    plan = build_write_plan(source, patches)
    return plan, original_cell, replacement_cell, original_before, replacement_before


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a safe operational replacement preview on a NEW .xlsm copy."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=(
            REPO_ROOT
            / "Excel"
            / "previews"
            / "Rehab_Center_System_v27_1_OPERATIONAL_PREVIEW.xlsm"
        ),
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    plan, original_cell, replacement_cell, original_before, replacement_before = (
        build_preview_plan(source)
    )
    report = apply_write_plan_to_copy(
        plan,
        output,
        overwrite=args.overwrite,
    )

    vba_preserved = _has_vba(source) == _has_vba(output) and _has_vba(output)

    print("OPERATIONAL PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(
        f"Original: {ORIGINAL_PROVIDER} {SLOT_TIME.strftime('%H:%M')} "
        f"-> THERAPIST_DAILY!{original_cell}"
    )
    print(
        f"Replacement: {REPLACEMENT_PROVIDER} {SLOT_TIME.strftime('%H:%M')} "
        f"-> THERAPIST_DAILY!{replacement_cell}"
    )
    print(f"Patient: {PATIENT_NAME}")
    print(f"Original cell before: {original_before!r}")
    print(f"Replacement cell before: {replacement_before!r}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print(
        "NEXT: open ONLY the operational preview. Confirm the original patient is "
        "muted/italic (NOT struck through), the line below reads → Φιλιππούσης 12:15, "
        "and the same patient appears under Φιλιππούσης at 12:15."
    )
    return 0 if report.source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
