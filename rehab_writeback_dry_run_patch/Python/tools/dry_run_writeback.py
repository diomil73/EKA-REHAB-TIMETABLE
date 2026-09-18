from __future__ import annotations

import argparse
from pathlib import Path

from rehab_excel.writeback import CellPatch, build_write_plan, dry_run_writeback


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate a future Excel write without modifying the workbook."
    )
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--sheet", default="THERAPIST_DAILY")
    parser.add_argument("--cell", default="A1")
    parser.add_argument("--value", default="DRY RUN")
    args = parser.parse_args()

    plan = build_write_plan(
        args.workbook,
        [CellPatch(sheet=args.sheet, cell=args.cell, value=args.value)],
    )
    report = dry_run_writeback(plan)

    print(f"Workbook: {report.workbook_path}")
    print(f"Operations: {report.operation_count}")
    print(f"Sheets: {', '.join(report.touched_sheets)}")
    print(f"Unchanged: {report.workbook_unchanged}")
    print(f"SHA256: {report.sha256_after}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
