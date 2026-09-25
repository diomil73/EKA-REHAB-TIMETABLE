from __future__ import annotations

import argparse
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.registration import NewTherapistRequest  # noqa: E402
from rehab_excel.therapist_registration import (  # noqa: E402
    TherapistRegistrationWriteError,
    create_therapist_registration_preview,
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
            "Create a safe preview copy with one validated physiotherapist appended "
            "to SETTINGS!THERAPISTS_FTH."
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
        default=REPO_ROOT / "Excel" / "previews" / "NEW_THERAPIST_PREVIEW.xlsm",
    )
    parser.add_argument("--name", required=True)
    parser.add_argument("--robotic-capable", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    request = NewTherapistRequest(
        display_name=args.name,
        robotic_capable=args.robotic_capable,
    )

    try:
        report = create_therapist_registration_preview(
            args.source,
            args.output,
            request,
            overwrite=args.overwrite,
        )
    except TherapistRegistrationWriteError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    output = Path(report.output_path)
    vba_present = _has_vba(output)

    print("NEW THERAPIST PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Therapist: {report.therapist_name}")
    print(f"SETTINGS row: {report.excel_row}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Read-back verified: {report.verified_in_output}")
    print(f"Other SETTINGS values unchanged: {report.other_settings_unchanged}")
    print(f"VBA project present: {vba_present}")
    print("NEXT: open only NEW_THERAPIST_PREVIEW.xlsm and inspect SETTINGS column A.")
    return (
        0
        if report.source_unchanged
        and report.verified_in_output
        and report.other_settings_unchanged
        and vba_present
        else 3
    )


if __name__ == "__main__":
    raise SystemExit(main())
