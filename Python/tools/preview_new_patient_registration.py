from __future__ import annotations

import argparse
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.models import PatientType  # noqa: E402
from rehab_core.registration import NewPatientRequest  # noqa: E402
from rehab_excel.patient_registration import (  # noqa: E402
    PatientRegistrationWriteError,
    create_patient_registration_preview,
)


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create a safe preview copy with one formally validated new patient "
            "appended to the authoritative PATIENTS sheet. PATIENT_PLANNER is not changed."
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
        default=REPO_ROOT / "Excel" / "previews" / "NEW_PATIENT_PREVIEW.xlsm",
    )
    parser.add_argument("--patient-id", required=True)
    parser.add_argument("--name", required=True)
    parser.add_argument("--room")
    parser.add_argument("--status")
    parser.add_argument("--infectious", action="store_true")
    parser.add_argument(
        "--patient-type",
        choices=(PatientType.INPATIENT.value, PatientType.OUTPATIENT.value),
        default=PatientType.INPATIENT.value,
        help="Patient classification; defaults to inpatient for backward compatibility",
    )
    parser.add_argument("--hospital-mrn")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    request = NewPatientRequest(
        patient_id=args.patient_id,
        display_name=args.name,
        room=args.room,
        infectious=args.infectious,
        status=args.status,
        patient_type=PatientType(args.patient_type),
        hospital_mrn=args.hospital_mrn,
    )

    try:
        report = create_patient_registration_preview(
            args.source,
            args.output,
            request,
            overwrite=args.overwrite,
        )
    except (PatientRegistrationWriteError, ValueError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    source = Path(report.source_path)
    output = Path(report.output_path)
    vba_present = _has_vba(output)

    print("NEW PATIENT PREVIEW OK")
    print(f"Source: {source}")
    print(f"Preview: {output}")
    print(f"PatientID: {report.patient_id}")
    print(f"Patient type: {args.patient_type}")
    print(f"Hospital MRN: {args.hospital_mrn or ''}")
    print(f"PATIENTS row: {report.excel_row}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Read-back verified: {report.verified_in_output}")
    print(f"VBA project present: {vba_present}")
    print("PATIENT_PLANNER changed: False")
    print(f"NEXT: open only {output.name} and inspect the appended PATIENTS row.")
    return 0 if report.source_unchanged and report.verified_in_output and vba_present else 3


if __name__ == "__main__":
    raise SystemExit(main())
