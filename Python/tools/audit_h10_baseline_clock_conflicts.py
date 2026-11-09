from __future__ import annotations

"""Read-only clock-level conflict audit of existing finalized therapist history.

Displays no patient names/IDs or session IDs. Does not modify workbooks.
Use to diagnose legacy, non-synthetic baseline facts separately from H8 demos.
"""

import argparse
from collections import Counter, defaultdict
from datetime import date
from hashlib import sha256
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_core.daily_state import DailySessionStatus
from rehab_excel.patient_session_history_store import read_current_patient_session_history_rows
from tools.report_therapist_history import report_therapist_history


def clock_conflicts(reports, sessions, target_date: date) -> tuple[dict, ...]:
    """Summarize actual delivered collisions, claimed capacity and approvals."""
    day_reports = [r for r in reports if r.service_date == target_date]
    if not day_reports:
        return ()
    active = {DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED}
    clock_map = defaultdict(Counter)
    for session in sessions:
        if session.service_date != target_date or session.status not in active:
            continue
        if not session.effective_therapist_id or session.effective_time is None:
            raise ValueError("Delivered session lacks effective therapist/time")
        clock_map[(session.snapshot_id, session.effective_therapist_id.casefold())][
            session.effective_time.strftime("%H:%M:%S")
        ] += 1
    findings = []
    for row in day_reports:
        counts = clock_map[(row.snapshot_id, row.therapist_id.casefold())]
        collisions = tuple(sorted((clock, n) for clock, n in counts.items() if n > 1))
        distinct = len(counts)
        delivered = sum(counts.values())
        problems = []
        if row.capacity_timeslots is None or row.capacity_timeslots <= 0:
            problems.append("CAPACITY_NOT_CONFIGURED")
        elif row.final_occupied_timeslots > row.capacity_timeslots or distinct > row.capacity_timeslots:
            problems.append("OCCUPIED_SLOTS_EXCEED_CAPACITY")
        if row.planned_occupied_timeslots > (row.capacity_timeslots or 0):
            problems.append("INITIAL_SLOTS_EXCEED_CAPACITY")
        if collisions and row.approved_shared_or_over_capacity_events == 0:
            problems.append("SHARED_CLOCK_WITHOUT_RECORDED_APPROVAL")
        if distinct != row.final_occupied_timeslots or delivered != row.final_active_sessions:
            problems.append("SLOT_OR_SESSION_RECONCILIATION_MISMATCH")
        if problems:
            findings.append({
                "therapist": row.therapist_id,
                "snapshot": row.snapshot_id,
                "capacity": row.capacity_timeslots,
                "planned_slots": row.planned_occupied_timeslots,
                "final_slots": row.final_occupied_timeslots,
                "delivered_sessions": delivered,
                "distinct_clock_times": distinct,
                "approved_shared": row.approved_shared_or_over_capacity_events,
                "collisions": collisions,
                "flags": tuple(problems),
            })
    return tuple(sorted(findings, key=lambda f: f["therapist"].casefold()))


def audit(path: Path, target_date: date) -> int:
    if not path.is_file() or path.suffix.casefold() != ".xlsm":
        print("SAFETY STOP: existing XLSM source required")
        return 2
    before = sha256(path.read_bytes()).hexdigest()
    try:
        findings = clock_conflicts(
            report_therapist_history(path),
            read_current_patient_session_history_rows(path),
            target_date,
        )
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    if sha256(path.read_bytes()).hexdigest() != before:
        print("SAFETY STOP: Workbook changed during read-only conflict audit")
        return 2
    print(f"BASELINE CLOCK CONFLICT AUDIT {target_date.isoformat()}")
    for row in findings:
        collision_text = ",".join(f"{t}({count})" for t, count in row["collisions"])
        print(
            f"{row['therapist']} | capacity={row['capacity']} | "
            f"initial_slots={row['planned_slots']} | final_slots={row['final_slots']} | "
            f"delivered={row['delivered_sessions']} | distinct_times={row['distinct_clock_times']} | "
            f"approved_shared={row['approved_shared']} | overlapping={collision_text or 'none'} | "
            f"flags={','.join(row['flags'])}"
        )
    print("Workbook unchanged: True")
    if findings:
        print(f"REVIEW REQUIRED: {len(findings)} therapist/day exceptions (not an acceptance proof)")
    else:
        print("NO CLOCK/APPROVAL EXCEPTIONS DETECTED for requested day")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only finalized clock conflict audit, no patient identifiers")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--date", type=date.fromisoformat, default=date(2026, 10, 30))
    args = parser.parse_args()
    return audit(args.workbook, args.date)


if __name__ == "__main__":
    raise SystemExit(main())
