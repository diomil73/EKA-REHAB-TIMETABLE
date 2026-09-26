from __future__ import annotations

import argparse
from datetime import datetime, time
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.outpatient_schedule_registration import (  # noqa: E402
    OutpatientScheduleRequest,
    OutpatientScheduleWriteError,
    create_outpatient_schedule_preview,
)


def _parse_time(value: str) -> time:
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Time must be HH:MM") from exc


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create a safe preview copy with one outpatient recurring schedule row "
            "appended or updated in OUTPATIENT_SCHEDULE."
        )
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--patient-id", required=True)
    parser.add_argument("--treatment", required=True)
    parser.add_argument("--time", type=_parse_time, required=True)
    parser.add_argument("--days", required=True)
    parser.add_argument("--therapist")
    parser.add_argument("--target-base-entry-id")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    request = OutpatientScheduleRequest(
        patient_id=args.patient_id,
        treatment=args.treatment,
        start_time=args.time,
        day_pattern=args.days,
        therapist_id=args.therapist,
        target_base_entry_id=args.target_base_entry_id,
    )

    try:
        report = create_outpatient_schedule_preview(
            args.source,
            args.output,
            request,
            overwrite=args.overwrite,
        )
    except OutpatientScheduleWriteError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    output = Path(report.output_path)
    vba_present = _has_vba(output)

    print("OUTPATIENT SCHEDULE PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Base entry ID: {report.base_entry_id}")
    print(f"OUTPATIENT_SCHEDULE row: {report.excel_row}")
    print(f"Created: {report.created}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Read-back verified: {report.verified_in_output}")
    print(f"VBA project present: {vba_present}")
    print("PATIENT_PLANNER changed: False")
    print(f"NEXT: open only {output.name} and inspect OUTPATIENT_SCHEDULE row {report.excel_row}.")
    return 0 if report.source_unchanged and report.verified_in_output and vba_present else 3


if __name__ == "__main__":
    raise SystemExit(main())
