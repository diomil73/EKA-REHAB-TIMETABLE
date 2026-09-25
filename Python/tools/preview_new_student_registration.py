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

from rehab_core.registration import NewStudentRequest  # noqa: E402
from rehab_excel.student_registration import (  # noqa: E402
    StudentRegistrationWriteError,
    create_student_registration_preview,
)
from rehab_excel.student_registry_schema import StudentRegistrySchemaError  # noqa: E402


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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a safe preview copy with one validated student registered."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "NEW_STUDENT_PREVIEW.xlsm",
    )
    parser.add_argument("--student-id", required=True)
    parser.add_argument("--name", required=True)
    parser.add_argument("--number", required=True, type=int)
    parser.add_argument("--start", required=True, type=_parse_date)
    parser.add_argument("--end", required=True, type=_parse_date)
    parser.add_argument("--supervisor")
    parser.add_argument(
        "--no-replacements",
        action="store_true",
        help="Register the student as not eligible for replacements.",
    )
    parser.add_argument("--robotic-capable", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    request = NewStudentRequest(
        student_id=args.student_id,
        display_name=args.name,
        student_number=args.number,
        placement_start=args.start,
        placement_end=args.end,
        supervisor_therapist_id=args.supervisor,
        replacement_capable=not args.no_replacements,
        robotic_capable=args.robotic_capable,
    )

    try:
        report = create_student_registration_preview(
            args.source,
            args.output,
            request,
            overwrite=args.overwrite,
        )
    except (StudentRegistrationWriteError, StudentRegistrySchemaError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    output = Path(report.output_path)
    vba_present = _has_vba(output)

    print("NEW STUDENT PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"StudentID: {report.student_id}")
    print(f"STUDENTS row: {report.excel_row}")
    print(f"Schema created: {report.schema_created}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Read-back verified: {report.verified_in_output}")
    print(f"VBA project present: {vba_present}")
    print("NEXT: open only NEW_STUDENT_PREVIEW.xlsm and inspect STUDENTS.")
    return (
        0
        if report.source_unchanged and report.verified_in_output and vba_present
        else 3
    )


if __name__ == "__main__":
    raise SystemExit(main())
