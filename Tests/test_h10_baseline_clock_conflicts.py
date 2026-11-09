from datetime import date, time
from types import SimpleNamespace

from rehab_core.daily_state import DailySessionStatus
from tools.audit_h10_baseline_clock_conflicts import clock_conflicts


DAY = date(2026, 10, 30)


def report(therapist, capacity, planned, final, delivered, approved=0):
    return SimpleNamespace(
        service_date=DAY, snapshot_id="finalized:2026-10-30:v2",
        therapist_id=therapist, capacity_timeslots=capacity,
        planned_occupied_timeslots=planned, final_occupied_timeslots=final,
        final_active_sessions=delivered,
        approved_shared_or_over_capacity_events=approved,
    )


def session(therapist, clock):
    return SimpleNamespace(
        service_date=DAY, snapshot_id="finalized:2026-10-30:v2",
        effective_therapist_id=therapist,
        effective_time=time.fromisoformat(clock), status=DailySessionStatus.ACTIVE,
    )


def test_baseline_oct30_flags_actual_seventh_clock_and_duplicated_times():
    rows = [
        report("Σκαρώνης", 6, 7, 7, 7),
        report("Σάρρας", 6, 6, 6, 7),
        report("Χρήστου", 6, 5, 5, 6),
    ]
    times = ("08:30", "09:15", "10:00", "10:45", "11:30", "12:15")
    sessions = (
        [session("Σκαρώνης", t) for t in (*times, "13:00")] +
        [session("Σάρρας", t) for t in (*times, "12:15")] +
        [session("Χρήστου", t) for t in (*times[:5], "10:45")]
    )
    findings = {item["therapist"]: item for item in clock_conflicts(rows, sessions, DAY)}
    assert "OCCUPIED_SLOTS_EXCEED_CAPACITY" in findings["Σκαρώνης"]["flags"]
    assert "INITIAL_SLOTS_EXCEED_CAPACITY" in findings["Σκαρώνης"]["flags"]
    assert findings["Σάρρας"]["collisions"] == (("12:15:00", 2),)
    assert findings["Χρήστου"]["collisions"] == (("10:45:00", 2),)
    assert all(
        "SHARED_CLOCK_WITHOUT_RECORDED_APPROVAL" in findings[name]["flags"]
        for name in ("Σάρρας", "Χρήστου")
    )


def test_approved_overlap_is_not_mislabeled_unapproved():
    rows = [report("T1", 6, 1, 1, 2, approved=1)]
    sessions = [session("T1", "08:30"), session("T1", "08:30")]
    assert clock_conflicts(rows, sessions, DAY) == ()


def test_clean_day_has_no_findings():
    rows = [report("T1", 6, 1, 1, 1)]
    assert clock_conflicts(rows, [session("T1", "08:30")], DAY) == ()
