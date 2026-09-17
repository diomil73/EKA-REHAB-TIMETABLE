from __future__ import annotations

import argparse
from pathlib import Path

from rehab_excel import audit_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only Rehab workbook audit")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    audit = audit_workbook(args.workbook)
    print(f"Workbook: {audit.workbook_path}")
    print(f"Patients: {audit.patient_count}")
    print(f"Importable base entries: {audit.base_entry_count}")
    print(f"Provider-owned entries: {audit.provider_entry_count}")
    if not audit.issues:
        print("Issues: none")
    else:
        print("Issues:")
        for issue in audit.issues:
            print(
                f"  [{issue.severity.upper()}] {issue.code}: "
                f"{issue.message} (count={issue.count})"
            )
    return 0 if audit.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
