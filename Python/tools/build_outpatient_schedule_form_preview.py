from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.outpatient_schedule_form_vba import (  # noqa: E402
    OutpatientScheduleFormVbaError,
    create_outpatient_schedule_form_preview,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a safe .xlsm preview with the outpatient schedule UserForm.")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        report = create_outpatient_schedule_form_preview(args.source, args.output, overwrite=args.overwrite)
    except OutpatientScheduleFormVbaError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("OUTPATIENT SCHEDULE FORM PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA project present: {report.vba_present}")
    print(f"Form present: {report.form_present}")
    print(f"Run macro: {report.macro_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
