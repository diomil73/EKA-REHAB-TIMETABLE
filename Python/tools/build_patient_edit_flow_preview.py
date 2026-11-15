from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.patient_edit_flow_vba import (  # noqa: E402
    PatientEditFlowVbaError,
    create_patient_edit_flow_preview,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Add patient edit form and registry-menu entry to a safe XLSM copy.")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        report = create_patient_edit_flow_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except PatientEditFlowVbaError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("PATIENT EDIT FLOW PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Edit form present: {report.edit_form_present}")
    print(f"Edit button present: {report.edit_button_present}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
