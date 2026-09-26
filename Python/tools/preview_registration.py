from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.registration import (  # noqa: E402
    NewPatientRequest,
    NewStudentRequest,
    NewTherapistRequest,
)
from rehab_excel.registration_orchestrator import (  # noqa: E402
    RegistrationOrchestrationError,
    create_registration_preview,
)


def _parse_date(value: str):
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            pass
    raise argparse.ArgumentTypeError(
        f"Unsupported date {value!r}; use DD/MM/YYYY or YYYY-MM-DD"
    )


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Unified safe registration preview for patient, therapist, or student."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--preview-dir",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews",
    )
    parser.add_argument("--overwrite", action="store_true")

    subparsers = parser.add_subparsers(dest="entity", required=True)

    patient = subparsers.add_parser("patient", help="Register a patient in preview.")
    patient.add_argument("--patient-id", required=True)
    patient.add_argument("--name", required=True)
    patient.add_argument("--room")
    patient.add_argument("--infectious", action="store_true")
    patient.add_argument("--status")

    therapist = subparsers.add_parser("therapist", help="Register a therapist in preview.")
    therapist.add_argument("--name", required=True)
    therapist.add_argument("--robotic-capable", action="store_true")

    student = subparsers.add_parser("student", help="Register a student in preview.")
    student.add_argument("--student-id", required=True)
    student.add_argument("--name", required=True)
    student.add_argument("--number", required=True, type=int)
    student.add_argument("--start", required=True, type=_parse_date)
    student.add_argument("--end", required=True, type=_parse_date)
    student.add_argument("--supervisor")
    student.add_argument("--no-replacements", action="store_true")
    student.add_argument("--robotic-capable", action="store_true")

    return parser


def _request_from_args(args):
    if args.entity == "patient":
        return NewPatientRequest(
            patient_id=args.patient_id,
            display_name=args.name,
            room=args.room,
            infectious=args.infectious,
            status=args.status,
        )
    if args.entity == "therapist":
        return NewTherapistRequest(
            display_name=args.name,
            robotic_capable=args.robotic_capable,
        )
    return NewStudentRequest(
        student_id=args.student_id,
        display_name=args.name,
        student_number=args.number,
        placement_start=args.start,
        placement_end=args.end,
        supervisor_therapist_id=args.supervisor,
        replacement_capable=not args.no_replacements,
        robotic_capable=args.robotic_capable,
    )


def main() -> int:
    args = _build_parser().parse_args()
    request = _request_from_args(args)

    try:
        result = create_registration_preview(
            args.source,
            args.preview_dir,
            request,
            overwrite=args.overwrite,
        )
    except RegistrationOrchestrationError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    output = Path(result.output_path)
    vba_present = _has_vba(output)

    print("REGISTRATION PREVIEW OK")
    print(f"Entity: {result.kind.value}")
    print(f"Subject: {result.subject_key}")
    print(f"Name: {result.display_name}")
    print(f"Preview: {result.output_path}")
    print(f"Excel row: {result.excel_row}")
    if result.schema_created is not None:
        print(f"Student schema created: {result.schema_created}")
    print(f"Source unchanged: {result.source_unchanged}")
    print(f"Read-back verified: {result.verified_in_output}")
    print(f"VBA project present: {vba_present}")
    return (
        0
        if result.source_unchanged and result.verified_in_output and vba_present
        else 3
    )


if __name__ == "__main__":
    raise SystemExit(main())
