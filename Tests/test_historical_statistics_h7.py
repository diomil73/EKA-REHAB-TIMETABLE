from datetime import date
from types import SimpleNamespace

import pytest

from rehab_core.daily_state import DailySessionStatus as Status
from rehab_core.day_finalization import FinalizedSnapshotStatus
from rehab_excel.historical_statistics import HistoricalStatisticsError, summarize_final_history


def fixture(version=2):
    d = date(2026, 10, 30)
    snap = SimpleNamespace(snapshot_id=f"finalized:2026-10-30:v{version}", service_date=d,
        version=version, status=FinalizedSnapshotStatus.CURRENT_FINAL, unresolved_sessions=0,
        planned_session_count=3, final_session_count=1, patient_absence_count=1,
        cancellation_count=1, replacement_count=0)
    patient = tuple(SimpleNamespace(snapshot_id=snap.snapshot_id, service_date=d,
        snapshot_version=version, session_id=f"s{i}", status=status)
        for i, status in enumerate((Status.ACTIVE, Status.PATIENT_ABSENT, Status.CANCELLED), 1))
    therapist = (SimpleNamespace(snapshot_id=snap.snapshot_id, service_date=d,
        therapist_id="t1", planned_sessions=3, final_active_sessions=1,
        replacements_received=0, released_by_cancellation_or_no_show=2,
        treatment_suspensions_affecting_load=0),)
    return snap, patient, therapist


def test_daily_and_therapist_aggregate():
    snap, patients, therapists = fixture()
    days, people = summarize_final_history((snap,), patients, therapists)
    assert len(days) == 1 and days[0].version == 2
    assert (days[0].scheduled, days[0].performed, days[0].patient_absences, days[0].cancellations) == (3, 1, 1, 1)
    assert len(people) == 1 and people[0].performed_sessions == 1
    assert people[0].days == 1


def test_outside_period_does_not_count():
    snap, patients, therapists = fixture()
    days, people = summarize_final_history((snap,), patients, therapists,
        start=date(2026, 11, 1))
    assert days == () and people == ()


def test_current_snapshot_disagreement_is_safety_stop():
    snap, patients, therapists = fixture()
    snap.final_session_count = 2
    with pytest.raises(HistoricalStatisticsError, match="Snapshot totals disagree"):
        summarize_final_history((snap,), patients, therapists)


def test_missing_patient_row_is_safety_stop():
    snap, patients, therapists = fixture()
    with pytest.raises(HistoricalStatisticsError, match="Missing patient history"):
        summarize_final_history((snap,), patients[:-1], therapists)


def test_duplicate_current_date_is_safety_stop():
    snap, patients, therapists = fixture()
    with pytest.raises(HistoricalStatisticsError, match="Duplicate current"):
        summarize_final_history((snap, snap), patients, therapists)


def test_superseded_snapshot_rejected():
    snap, patients, therapists = fixture()
    snap.status = FinalizedSnapshotStatus.SUPERSEDED
    with pytest.raises(HistoricalStatisticsError, match="Non-current"):
        summarize_final_history((snap,), patients, therapists)


def test_invalid_dates_rejected():
    with pytest.raises(HistoricalStatisticsError, match="start date"):
        summarize_final_history((), (), (), start=date(2026, 11, 1), end=date(2026, 10, 1))
