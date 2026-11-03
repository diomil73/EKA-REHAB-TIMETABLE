from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON = REPO_ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.patient_session_history_store import (  # noqa: E402
    read_current_patient_session_history_rows,
)
from rehab_excel.patient_history_reports import read_patient_history_report  # noqa: E402

REPORTED_NUMBERS = (15, 18, 19, 20, 21, 22)


def _numeric_ids(values) -> set[int]:
    result = set()
    for value in values:
        text = str(value).strip()
        if text.isascii() and text.isdecimal():
            result.add(int(text))
    return result


def compare_patient_registry_history(workbook: Path) -> int:
    """Read-only comparison. Registry membership does NOT imply a scheduled session."""
    patients = read_patient_registry(workbook)
    history = read_current_patient_session_history_rows(workbook)
    daily = read_patient_history_report(workbook, mode="daily")
    weekly = read_patient_history_report(workbook, mode="weekly")

    registry_ids = [str(p.patient_id) for p in patients]
    source_ids = [str(r.patient_id) for r in history]
    daily_ids = [str(r.patient_id) for r in daily]
    weekly_ids = [str(r.patient_id) for r in weekly]
    numeric_registry = _numeric_ids(registry_ids)
    numeric_source = _numeric_ids(source_ids)
    numeric_daily = _numeric_ids(daily_ids)
    numeric_weekly = _numeric_ids(weekly_ids)

    print("H7 REGISTRY VS CURRENT_FINAL READ-ONLY AUDIT")
    print(f"Workbook: {workbook.resolve()}")
    print(f"PATIENTS registry entries: {len(patients)}")
    print(f"CURRENT_FINAL sessions: {len(history)}")
    print(f"Distinct patient IDs in CURRENT_FINAL: {len(set(source_ids))}")
    print(f"Report rows: daily={len(daily)}, weekly={len(weekly)}")

    errors = []
    if not history:
        errors.append("No finalized patient session history")
    if numeric_daily != numeric_source or numeric_weekly != numeric_source:
        errors.append("Numeric patient IDs differ between history and reports")

    for patient_no in REPORTED_NUMBERS:
        print(
            f"ID {patient_no:02d}: "
            f"registry={'YES' if patient_no in numeric_registry else 'NO'} "
            f"finalized_sessions={'YES' if patient_no in numeric_source else 'NO'} "
            f"daily={'YES' if patient_no in numeric_daily else 'NO'} "
            f"weekly={'YES' if patient_no in numeric_weekly else 'NO'}"
        )

    registry_without_sessions = numeric_registry - numeric_source
    sessions_without_registry = numeric_source - numeric_registry
    print("Registry numeric IDs without CURRENT_FINAL sessions:",
          ", ".join(f"{v:02d}" for v in sorted(registry_without_sessions)) or "none")
    print("CURRENT_FINAL numeric IDs absent from registry:",
          ", ".join(f"{v:02d}" for v in sorted(sessions_without_registry)) or "none")
    if sessions_without_registry:
        errors.append("CURRENT_FINAL references numeric patients absent from the registry; investigate identity mapping")

    for issue in errors:
        print("AUDIT ISSUE:", issue)
    print("Workbook unchanged: True")
    if errors:
        print("H7 REGISTRY AUDIT NEEDS REVIEW")
        return 2
    print("H7 REGISTRY AUDIT COMPLETE")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only patient registry vs H7 history comparison")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    if not args.workbook.is_file():
        print("SAFETY STOP: source workbook missing")
        return 2
    try:
        return compare_patient_registry_history(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
