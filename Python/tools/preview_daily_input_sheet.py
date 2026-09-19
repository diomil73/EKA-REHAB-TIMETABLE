from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_input_patients import audit_daily_input_patients  # noqa: E402
from rehab_excel.daily_input_sheet import (  # noqa: E402
    DailyInputSheetError,
    build_daily_input_spec,
    create_daily_input_preview,
)
from rehab_excel.reader import read_settings  # noqa: E402


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a safe XLSM preview with a compact DAILY_INPUT sheet."
    )
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "DAILY_INPUT_PREVIEW.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    try:
        patient_audit = audit_daily_input_patients(source)
        if not patient_audit.ok:
            details = list(patient_audit.identity_mismatches) + [
                f"Planner-only patient: {name}" for name in patient_audit.planner_without_patient
            ]
            raise ValueError("Patient identity audit failed: " + "; ".join(details))

        settings = read_settings(source)
        spec = build_daily_input_spec(
            target_date=args.date,
            therapist_names=settings.therapist_names,
            patient_names=patient_audit.dropdown_names,
            patient_statuses=settings.patient_statuses,
            timeslots=settings.standard_timeslots,
        )
        preview, source_unchanged = create_daily_input_preview(
            source,
            output,
            spec,
            overwrite=args.overwrite,
        )
    except (ValueError, DailyInputSheetError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    vba_preserved = _has_vba(source) == _has_vba(preview) and _has_vba(preview)
    print("DAILY INPUT PREVIEW OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Source: {source}")
    print(f"Preview: {preview}")
    print(f"Therapists in dropdown: {len(spec.therapist_names)}")
    print(f"PATIENTS identities: {patient_audit.patient_rows}")
    print(f"PATIENT_PLANNER identities: {patient_audit.planner_rows}")
    print(f"Patients in dropdown: {len(spec.patient_names)}")
    if patient_audit.patients_without_planner:
        print(
            "PATIENTS rows without planner identity (kept in dropdown): "
            + ", ".join(patient_audit.patients_without_planner)
        )
    print(f"Patient status options: {len(spec.patient_statuses)}")
    print(f"Timeslots: {', '.join(spec.timeslots)}")
    print(f"Source unchanged: {source_unchanged}")
    print(f"VBA preserved: {vba_preserved}")
    print("NEXT: inspect the cleaned DAILY_INPUT sheet and its dropdowns.")
    return 0 if source_unchanged and vba_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
