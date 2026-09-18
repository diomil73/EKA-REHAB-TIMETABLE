from __future__ import annotations

import argparse
from pathlib import Path

from rehab_excel.layout_contract import audit_layout
from rehab_excel.package_probe import WorkbookPackageProbe


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit the confirmed Excel cell/layout contract without modifying the workbook."
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    report = audit_layout(args.workbook)
    probe = WorkbookPackageProbe(args.workbook)

    print(f"Workbook: {args.workbook}")
    for sheet in (
        "MASTER_SCHEDULE",
        "THERAPIST_DAILY",
        "REPLACEMENTS",
        "REPLACEMENT_LOG",
        "CONFLICT_LOG",
        "THERAPIST_ATTENDANCE",
    ):
        if sheet in probe.sheet_names:
            info = probe.sheet_info(sheet)
            print(f"{sheet}: {info.dimension or '(no dimension)'}")

    if not report.issues:
        print("Layout audit: OK")
        return 0

    for issue in report.issues:
        print(f"[{issue.severity.upper()}] {issue.code}: {issue.message}")

    print("Layout audit:", "OK WITH WARNINGS" if report.ok else "FAILED")
    return 0 if report.ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
