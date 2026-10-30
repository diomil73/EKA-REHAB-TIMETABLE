from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.outpatient_schedule_edit_vba import (  # noqa: E402
    EDIT_MACRO_NAME,
    install_outpatient_schedule_edit_support_in_workbook,
)
from rehab_excel.outpatient_schedule_form_vba import (  # noqa: E402
    OutpatientScheduleFormVbaError,
    create_outpatient_schedule_form_preview,
)
from rehab_excel.outpatient_schedule_manager_vba import (  # noqa: E402
    install_outpatient_schedule_manager_in_workbook,
)
from rehab_excel.outpatient_schedule_sheet import (  # noqa: E402
    OutpatientScheduleSheetError,
    ensure_outpatient_schedule_sheet_in_workbook,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a safe .xlsm preview with the outpatient schedule UserForm.")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        report = create_outpatient_schedule_form_preview(args.source, args.output, overwrite=args.overwrite)
        ensure_outpatient_schedule_sheet_in_workbook(report.output_path)
        install_outpatient_schedule_edit_support_in_workbook(report.output_path)
        install_outpatient_schedule_manager_in_workbook(report.output_path)
    except (OutpatientScheduleFormVbaError, OutpatientScheduleSheetError, RuntimeError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("OUTPATIENT SCHEDULE FORM PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA project present: {report.vba_present}")
    print(f"Form present: {report.form_present}")
    print(f"Create macro: {report.macro_name}")
    print(f"Edit macro: {EDIT_MACRO_NAME}")
    print("Outpatient schedule source: verified")
    print("Outpatient schedule manager: installed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
