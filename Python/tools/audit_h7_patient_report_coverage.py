from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_core.daily_state import DailySessionStatus  # noqa: E402
from rehab_excel.patient_history_reports import read_patient_history_report  # noqa: E402
from rehab_excel.patient_session_history_store import (  # noqa: E402
    read_current_patient_session_history_rows,
)
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots  # noqa: E402


def _numeric_gap_report(ids: set[str]) -> tuple[str, ...]:
    numeric = {int(value) for value in ids if value.isascii() and value.isdecimal()}
    if not numeric:
        return ()
    return tuple(f"{number:02d}" for number in range(min(numeric), max(numeric) + 1) if number not in numeric)


def audit_patient_report_coverage(workbook: Path) -> int:
    """Read-only source-to-report reconciliation, without inventing patient visits."""
    snapshots = read_finalized_day_snapshots(workbook, current_only=True)
    rows = read_current_patient_session_history_rows(workbook)
    print("H7 PATIENT COVERAGE READ-ONLY AUDIT")
    print(f"Workbook: {workbook.resolve()}")

    if not snapshots:
        print("SAFETY STOP: no CURRENT_FINAL history")
        return 2

    daily_rows = defaultdict(list)
    for row in rows:
        daily_rows[row.service_date].append(row)

    errors = []
    for snap in sorted(snapshots, key=lambda item: item.service_date):
        sessions = daily_rows[snap.service_date]
        performed = sum(r.status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED) for r in sessions)
        absent = sum(r.status == DailySessionStatus.PATIENT_ABSENT for r in sessions)
        cancelled = sum(r.status == DailySessionStatus.CANCELLED for r in sessions)
        replaced = sum(r.status == DailySessionStatus.REPLACED for r in sessions)
        actual = (performed, absent, cancelled, replaced)
        expected = (
            snap.final_session_count, snap.patient_absence_count,
            snap.cancellation_count, snap.replacement_count,
        )
        if actual != expected or (snap.planned_session_count is not None and len(sessions) != snap.planned_session_count):
            errors.append(f"{snap.service_date}: snapshot/patient history mismatch")

        patient_ids = {str(r.patient_id) for r in sessions}
        missing_ids = _numeric_gap_report(patient_ids)
        print(
            f"DATE {snap.service_date}: sessions={len(sessions)}, "
            f"distinct_patients={len(patient_ids)}, "
            f"performed={performed}, absences={absent}, cancellations={cancelled}"
        )
        print("PATIENT 15 in source:", "YES" if any(v.isascii() and v.isdecimal() and int(v) == 15 for v in patient_ids) else "NO")
        print(
            "NUMERIC ID GAPS IN SOURCE (not proof of error):",
            ", ".join(missing_ids) if missing_ids else "none",
        )

    for mode in ("daily", "weekly", "monthly", "overall"):
        report = read_patient_history_report(workbook, mode=mode)
        report_totals = Counter()
        for item in report:
            report_totals["scheduled"] += item.scheduled
            report_totals["performed"] += item.performed
            report_totals["absent"] += item.patient_absences
            report_totals["cancelled"] += item.cancellations
            report_totals["replaced"] += item.replacements
        expected_totals = Counter(
            scheduled=len(rows),
            performed=sum(r.status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED) for r in rows),
            absent=sum(r.status == DailySessionStatus.PATIENT_ABSENT for r in rows),
            cancelled=sum(r.status == DailySessionStatus.CANCELLED for r in rows),
            replaced=sum(r.status == DailySessionStatus.REPLACED for r in rows),
        )
        print(
            f"MODE {mode}: report_rows={len(report)} "
            f"scheduled={report_totals['scheduled']} "
            f"performed={report_totals['performed']} "
            f"absences={report_totals['absent']} "
            f"cancellations={report_totals['cancelled']} "
            f"replacements={report_totals['replaced']}"
        )
        if report_totals != expected_totals:
            errors.append(f"{mode}: report totals differ from CURRENT_FINAL sessions")
        source_ids = {str(row.patient_id) for row in rows}
        report_ids = {str(item.patient_id) for item in report}
        if source_ids != report_ids:
            errors.append(f"{mode}: report omitted source patient identifiers")

    if errors:
        for issue in errors:
            print("SAFETY STOP:", issue)
        return 2
    print("H7 PATIENT COVERAGE AUDIT OK; no workbook changes")
    print("NOTE: Missing numbers in the source are informational; registry membership must be checked before zero-row presentation.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only current-history patient and report ID coverage")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    if not args.workbook.is_file():
        print("SAFETY STOP: workbook not found")
        return 2
    try:
        return audit_patient_report_coverage(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
