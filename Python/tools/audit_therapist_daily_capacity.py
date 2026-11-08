from __future__ import annotations

"""Read-only audit of therapist/day slot capacity vs sessions and patients.

These are deliberately three DIFFERENT quantities:
- final_occupied_timeslots: distinct occupied clock slots
- final_active_sessions: count of delivered sessions
- unique_patients: distinct patients over all delivered sessions that day

Multiple sessions/patients can share one time slot. Count > capacity is an
anomaly requiring investigation, NOT automatically proof of an extra slot.
"""

import argparse
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.report_therapist_history import report_therapist_history
from rehab_excel.patient_session_history_store import read_current_patient_session_history_rows
from tools.export_therapist_history_ui import _unique_patient_days


def capacity_anomalies(reports, unique_patients_per_day, therapist: str | None = None):
    person_filter = therapist.strip().casefold() if therapist else None
    checked = set()
    result = []
    for row in reports:
        key = (row.service_date, row.therapist_id.casefold())
        if key in checked:
            raise ValueError("Duplicate CURRENT_FINAL therapist/day")
        checked.add(key)
        if person_filter and row.therapist_id.casefold() != person_filter:
            continue
        capacity = row.capacity_timeslots
        occupied = int(row.final_occupied_timeslots)
        sessions = int(row.final_active_sessions)
        patients = int(unique_patients_per_day.get(key, 0))
        planned_occupied = int(row.planned_occupied_timeslots)
        if capacity is not None and int(capacity) < 0:
            raise ValueError("Negative slot capacity")
        if min(occupied, sessions, patients, planned_occupied) < 0:
            raise ValueError("Negative count in finalized row")
        flags = []
        if capacity is None or int(capacity) == 0:
            flags.append("CAPACITY_NOT_CONFIGURED")
        elif occupied > int(capacity):
            flags.append("OCCUPIED_SLOTS_EXCEED_CAPACITY")
        if sessions > occupied:
            flags.append("MULTIPLE_SESSIONS_PER_SLOT")
        if patients > occupied:
            flags.append("MULTIPLE_DISTINCT_PATIENTS_PER_SLOT")
        if sessions > int(row.planned_sessions):
            flags.append("DELIVERED_EXCEEDS_INITIAL_PLAN")
        # Include every flagged row, and all requested-person daily rows:
        # this makes actual vs planned clock slots legible without making
        # potentially legal 8 sessions into an asserted 8-slot violation.
        if flags or person_filter:
            result.append({
                "day": row.service_date.isoformat(),
                "therapist": row.therapist_id,
                "capacity_slots": capacity,
                "initial_occupied_slots": planned_occupied,
                "final_occupied_slots": occupied,
                "initial_sessions": int(row.planned_sessions),
                "delivered_sessions": sessions,
                "unique_patients": patients,
                "approved_shared_exceptions": int(
                    row.approved_shared_or_over_capacity_events
                ),
                "replacements_received": int(row.replacements_received),
                "flags": tuple(flags),
            })
    return sorted(result, key=lambda r: (r["day"], str(r["therapist"]).casefold()))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read-only finalized therapist capacity-vs-patients diagnostic"
    )
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--therapist", help="Exact therapist label, optional", default=None)
    args = parser.parse_args()
    if not args.workbook.is_file() or args.workbook.suffix.casefold() != ".xlsm":
        print("SAFETY STOP: .xlsm source required")
        return 2
    try:
        reports = report_therapist_history(args.workbook)
        patients = _unique_patient_days(
            read_current_patient_session_history_rows(args.workbook)
        )
        flagged = capacity_anomalies(reports, patients, args.therapist)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    print("THERAPIST CAPACITY/SESSIONS READ-ONLY AUDIT OK")
    print("NOTE: Sessions and unique patients can exceed occupied clock slots.")
    print(f"Finalized therapist/day records: {len(reports)}")
    print(f"Displayed days: {len(flagged)}")
    for row in flagged:
        print(
            f"{row['day']} | {row['therapist']} | capacity_slots={row['capacity_slots']}"
            f" | initial_slots={row['initial_occupied_slots']}"
            f" | final_slots={row['final_occupied_slots']}"
            f" | planned_sessions={row['initial_sessions']}"
            f" | delivered_sessions={row['delivered_sessions']}"
            f" | unique_patients={row['unique_patients']}"
            f" | shared_approved={row['approved_shared_exceptions']}"
            f" | replacements={row['replacements_received']}"
            f" | flags={','.join(row['flags']) or 'none'}"
        )
    print("No .xlsm modifications made")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
