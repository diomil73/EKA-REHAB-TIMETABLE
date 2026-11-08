from __future__ import annotations

"""Independent, read-only proof of a previously seeded H8 demo workbook.

No installer, save, VBA patch, worksheet modification or candidate retry.
Every synthetic CURRENT_FINAL date must reconcile with all three historical
data sources and the synthetic attendance ledger, not merely exist in Excel.
"""

import argparse
from collections import Counter, defaultdict
from datetime import date
import hashlib
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_core.daily_state import DailySessionStatus
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots
from rehab_excel.patient_session_history_store import read_patient_session_history_rows
from rehab_excel.therapist_day_history_store import read_therapist_day_history_rows
from rehab_excel.day_history_event_store import read_day_history_events
from rehab_excel.therapist_attendance_history_store import read_therapist_attendance_history_facts
from rehab_excel.patient_registry_source import read_patient_registry
from tools.seed_h8_synthetic_history import ATTENDANCE_SHEET, DEMO_SOURCE, _days


class DemoVerificationError(RuntimeError):
    """A demo workbook lacks verifiable, internally consistent history."""


def check(condition: bool, message: str) -> None:
    if not condition:
        raise DemoVerificationError(message)


def verify_demo_history(path: str | Path) -> dict[str, int]:
    path = Path(path)
    check(path.is_file() and path.suffix.casefold() == ".xlsm",
          "Existing .xlsm workbook required")
    check("_DEMO_" in path.name.upper(), "Refusing a non-demo workbook")
    sha_before = hashlib.sha256(path.read_bytes()).hexdigest()
    try:
        snapshots = read_finalized_day_snapshots(path)
        patients = read_patient_session_history_rows(path)
        therapist_rows = read_therapist_day_history_rows(path)
        events = read_day_history_events(path)
        attendances = read_therapist_attendance_history_facts(path)
        registry = read_patient_registry(path)

        demo_snapshots = [s for s in snapshots if s.snapshot_id.startswith("DEMO-H8-")]
        check(len(demo_snapshots) >= 12, "Fewer than 12 synthetic finalized days")
        check(all(s.status.value == "CURRENT_FINAL" and s.version == 1
                  and s.unresolved_sessions == 0 for s in demo_snapshots),
              "Synthetic history has non-current, amended or unresolved snapshot")
        days = {s.service_date for s in demo_snapshots}
        check(len(days) == len(demo_snapshots), "Duplicate synthetic CURRENT_FINAL date")
        check(all(s.service_date.weekday() < 5 for s in demo_snapshots),
              "Synthetic history includes weekend")
        check({d.strftime("%Y-%m") for d in days} == {"2026-09", "2026-10"},
              "Synthetic dates do not cover September and October 2026")
        check(all(s.snapshot_id == f"DEMO-H8-{s.service_date:%Y%m%d}-V1"
                  for s in demo_snapshots), "Synthetic snapshot ID/date mismatch")
        check(not any(s.service_date in days and not s.snapshot_id.startswith("DEMO-")
                      for s in snapshots),
              "Synthetic service date overlaps baseline snapshot history")

        ids = {s.snapshot_id for s in demo_snapshots}
        check(not any(r.snapshot_id.startswith("DEMO-") and r.snapshot_id not in ids
                      for r in patients), "Orphan synthetic patient-session row")
        check(not any(r.snapshot_id.startswith("DEMO-") and r.snapshot_id not in ids
                      for r in therapist_rows), "Orphan synthetic therapist-day row")
        check(not any(f.snapshot_id.startswith("DEMO-") and f.snapshot_id not in ids
                      for f in attendances), "Orphan synthetic attendance row")

        p_by_snapshot = defaultdict(list)
        for row in patients:
            if row.snapshot_id in ids:
                p_by_snapshot[row.snapshot_id].append(row)
        t_by_snapshot = defaultdict(list)
        for row in therapist_rows:
            if row.snapshot_id in ids:
                t_by_snapshot[row.snapshot_id].append(row)
        a_by_snapshot = defaultdict(list)
        for row in attendances:
            if row.snapshot_id in ids:
                a_by_snapshot[row.snapshot_id].append(row)
        e_by_snapshot = defaultdict(list)
        for event in events:
            if event.related_snapshot_id in ids:
                e_by_snapshot[event.related_snapshot_id].append(event)

        registered = {str(x.patient_id).strip() for x in registry}
        check(len(registered) >= 24, "Not enough source patients in registry")
        all_therapists: set[str] = set()
        for snap in demo_snapshots:
            sid, day = snap.snapshot_id, snap.service_date
            p_rows = p_by_snapshot[sid]
            t_rows = t_by_snapshot[sid]
            a_rows = a_by_snapshot[sid]
            e_rows = e_by_snapshot[sid]
            check(len(t_rows) == len(a_rows) == 6,
                  f"{day}: expected 6 therapist and attendance records")
            therapist_ids = {row.therapist_id for row in t_rows}
            check(len(therapist_ids) == 6, f"{day}: duplicate therapist/day")
            all_therapists.update(therapist_ids)
            check({row.therapist_id for row in a_rows} == therapist_ids,
                  f"{day}: attendance and workload therapist IDs differ")
            check(len(e_rows) == 1 and e_rows[0].event_type.value == "DAY_FINALIZED"
                  and e_rows[0].actor_or_source == DEMO_SOURCE
                  and e_rows[0].service_date == day,
                  f"{day}: missing synthetic finalization event")
            check(all(row.service_date == day and row.snapshot_version == 1
                      and row.patient_id in registered for row in p_rows),
                  f"{day}: patient sessions have wrong dates/versions/IDs")
            check(len({row.session_id for row in p_rows}) == len(p_rows)
                  and len({row.patient_id for row in p_rows}) == len(p_rows),
                  f"{day}: duplicate session or patient in synthetic day")
            count = Counter(row.status.value for row in p_rows)
            check(snap.planned_session_count == len(p_rows), f"{day}: planned count mismatch")
            check(snap.final_session_count == count["active"] + count["replaced"],
                  f"{day}: performed count mismatch")
            check(snap.patient_absence_count == count["patient_absent"],
                  f"{day}: patient absence mismatch")
            check(snap.cancellation_count == count["cancelled"],
                  f"{day}: cancellations mismatch")
            check(snap.no_show_count == 0 and snap.replacement_count == count["replaced"],
                  f"{day}: no-show/replacement mismatch")
            check(snap.treatment_suspension_count ==
                  sum(row.cancellation_kind == "treatment_suspended" for row in p_rows),
                  f"{day}: treatment suspension mismatch")
            check(snap.approved_exception_count == 0,
                  f"{day}: unexpected approved exception")
            check(all(row.status != DailySessionStatus.THERAPIST_ABSENT for row in p_rows),
                  f"{day}: unresolved therapist absence in finalized history")

            for trow in t_rows:
                tid = trow.therapist_id
                scheduled = [r for r in p_rows if r.original_therapist_id == tid]
                delivered = [r for r in p_rows if r.effective_therapist_id == tid
                             and r.status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED)]
                check(trow.service_date == day and trow.capacity_timeslots == 6,
                      f"{day}/{tid}: wrong workload date/capacity")
                check(trow.planned_sessions == len(scheduled)
                      and trow.planned_occupied_timeslots == len(scheduled),
                      f"{day}/{tid}: planned therapist totals mismatch")
                check(trow.final_active_sessions == len(delivered)
                      and trow.final_occupied_timeslots == len(delivered),
                      f"{day}/{tid}: delivered therapist totals mismatch")
                check(trow.replacements_received ==
                      sum(r.status == DailySessionStatus.REPLACED for r in delivered),
                      f"{day}/{tid}: replacement totals mismatch")
                check(trow.released_by_cancellation_or_no_show ==
                      sum(r.status == DailySessionStatus.PATIENT_ABSENT
                          or r.cancellation_kind == "department_postponed"
                          for r in scheduled),
                      f"{day}/{tid}: released-slot totals mismatch")
                check(trow.treatment_suspensions_affecting_load ==
                      sum(r.cancellation_kind == "treatment_suspended" for r in scheduled),
                      f"{day}/{tid}: suspension totals mismatch")
                check(trow.approved_shared_or_over_capacity_events == 0,
                      f"{day}/{tid}: unexpected capacity exception")
            for fact in a_rows:
                check(fact.service_date == day and fact.source == DEMO_SOURCE,
                      f"{day}: attendance provenance mismatch")
                delivered = sum(r.effective_therapist_id == fact.therapist_id
                                and r.status in (DailySessionStatus.ACTIVE,
                                                 DailySessionStatus.REPLACED) for r in p_rows)
                if fact.outcome.value == "ABSENT":
                    check(fact.reason_code == "SICK_LEAVE" and delivered == 0,
                          f"{day}/{fact.therapist_id}: absent but delivered sessions")
                else:
                    check(fact.outcome.value == "PRESENT" and not fact.reason_code,
                          f"{day}/{fact.therapist_id}: invalid attendance flag")
        check(len(all_therapists) == 6, "Therapist cohort changed across demo dates")
        check(len({fact.therapist_id for fact in attendances
                   if fact.snapshot_id in ids and fact.outcome.value == "ABSENT"}) >= 3,
              "Synthetic history lacks varied therapist absence examples")
        result = {
            "days": len(days),
            "patients": sum(map(len, p_by_snapshot.values())),
            "therapist_days": sum(map(len, t_by_snapshot.values())),
            "attendance_facts": sum(map(len, a_by_snapshot.values())),
            "finalization_events": sum(map(len, e_by_snapshot.values())),
        }
        return result
    finally:
        after = hashlib.sha256(path.read_bytes()).hexdigest()
        if sha_before != after:
            raise DemoVerificationError("Workbook changed during read-only verification")


def main() -> int:
    parser = argparse.ArgumentParser(description="Strict read-only synthetic history proof")
    parser.add_argument("--workbook", required=True, type=Path)
    args = parser.parse_args()
    try:
        result = verify_demo_history(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        print("Do not save, patch or retry seed in this candidate.")
        return 2
    print("H8 SYNTHETIC HISTORY CROSS-SHEET VERIFICATION OK")
    for name, amount in result.items():
        print(f"{name}: {amount}")
    print("Workbook unchanged: True")
    print("Source: SYNTHETIC DEMO ONLY, not real attendance history")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
