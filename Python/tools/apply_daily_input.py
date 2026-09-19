from __future__ import annotations

import argparse
import sys
from pathlib import Path
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_excel.daily_input_reader import DailyInputReadError, read_daily_input  # noqa: E402
from rehab_excel.native_excel import apply_write_plan_to_copy  # noqa: E402
from rehab_excel.patient_centric_preview import (  # noqa: E402
    PatientCentricPreviewError,
    build_patient_centric_preview_plan,
)
from rehab_excel.reader import read_base_schedule, read_patients  # noqa: E402


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Read DAILY_INPUT from a saved .xlsm and create a patient-centric "
            "operational timetable preview on a NEW copy."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "DAILY_INPUT_PREVIEW.xlsm",
        help="Saved workbook containing DAILY_INPUT entries",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "DAILY_INPUT_APPLIED_PREVIEW.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_book = args.input.resolve()
    output = args.output.resolve()
    if not input_book.exists():
        print(f"SAFETY STOP: input workbook not found: {input_book}")
        return 2

    try:
        # DAILY_INPUT date is read first only after loading patient/base data.
        patients = read_patients(input_book)
        base_entries = read_base_schedule(input_book)

        # Sessions are needed to resolve therapist labels. Read the date from
        # DAILY_INPUT with a temporary all-date materialization strategy: the
        # reader itself validates B2, then we materialize the exact date and
        # read once more with the correct dated sessions.
        from openpyxl import load_workbook
        from datetime import datetime, date

        wb = load_workbook(input_book, read_only=True, data_only=False, keep_vba=True)
        try:
            raw_date = wb["DAILY_INPUT"]["B2"].value
        finally:
            wb.close()
        if isinstance(raw_date, datetime):
            target_date = raw_date.date()
        elif isinstance(raw_date, date):
            target_date = raw_date
        else:
            target_date = datetime.strptime(str(raw_date).strip(), "%d/%m/%Y").date()

        sessions = materialize_sessions_for_date(base_entries, target_date)
        daily = read_daily_input(input_book, patients=patients, sessions=sessions)
        if daily.target_date != target_date:
            raise DailyInputReadError("DAILY_INPUT date changed during read")

        states = build_daily_session_states(
            sessions,
            absences=daily.absences,
            replacements=(),
            target_date=daily.target_date,
        )
        changed = [state for state in states if state.status.value != "active"]
        if not changed:
            raise DailyInputReadError(
                "DAILY_INPUT rows did not match any scheduled session on the selected date"
            )

        preview = build_patient_centric_preview_plan(
            input_book,
            base_entries=base_entries,
            states=states,
            sessions=sessions,
            patients=patients,
            rebuild_multi_member_groups=True,
        )
        report = apply_write_plan_to_copy(
            preview.write_plan,
            output,
            overwrite=args.overwrite,
        )
    except (DailyInputReadError, PatientCentricPreviewError, ValueError, KeyError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    vba_preserved = _has_vba(input_book) == _has_vba(output) and _has_vba(output)
    print("DAILY INPUT APPLIED PREVIEW OK")
    print(f"Date: {daily.target_date.isoformat()}")
    print(f"Input: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Therapist rows read: {daily.therapist_rows_used}")
    print(f"Patient rows read: {daily.patient_rows_used}")
    print(f"Operational absences: {len(daily.absences)}")
    print(f"Changed sessions: {len(preview.bindings)}")
    for warning in daily.warnings:
        print(f"WARNING: {warning}")
    for binding in preview.bindings:
        print(f"  {binding.status.value}: {binding.session_id} [{binding.original_cell}]")
    print(f"Input unchanged: {report.source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print("NEXT: open only DAILY_INPUT_APPLIED_PREVIEW.xlsm and inspect the affected patient line.")
    return 0 if report.source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
