from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.patient_registration_vertical_slice import (  # noqa: E402
    PatientRegistrationBridgeVbaError,
    create_patient_registration_bridge_preview,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Create a safe copy with frmNewPatient wired to the patient-centric in-session registration bridge."
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        report = create_patient_registration_bridge_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except PatientRegistrationBridgeVbaError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("PATIENT REGISTRATION BRIDGE PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA project present: {report.vba_present}")
    print(f"Patient form present: {report.form_present}")
    print(f"Bridge wired: {report.bridge_wired}")
    print("NEXT: open only the preview, run ShowRegistrationMenu, then submit a test patient.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
