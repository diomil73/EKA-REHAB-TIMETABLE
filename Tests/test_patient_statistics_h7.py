from datetime import date
from types import SimpleNamespace

import pytest

from rehab_core.daily_state import DailySessionStatus
from rehab_excel.patient_statistics import summarize_patient_sessions


def row(day, sid, pid, status):
    return SimpleNamespace(
        service_date=date.fromisoformat(day),
        session_id=sid,
        patient_id=pid,
        status=status,
    )


def test_patient_aggregation_performed_absent_cancelled_replaced():
    rows = (
        row("2026-10-30", "s1", "p1", DailySessionStatus.ACTIVE),
        row("2026-10-30", "s2", "p1", DailySessionStatus.PATIENT_ABSENT),
        row("2026-10-30", "s3", "p1", DailySessionStatus.REPLACED),
        row("2026-10-30", "s4", "p2", DailySessionStatus.CANCELLED),
    )
    stats = summarize_patient_sessions(rows)
    assert [(x.patient_id, x.scheduled, x.performed, x.absent, x.cancelled, x.replaced) for x in stats] == [
        ("p1", 3, 2, 1, 0, 1),
        ("p2", 1, 0, 0, 1, 0),
    ]


def test_date_filter_inclusive_and_invalid_order():
    rows = (row("2026-10-29", "s1", "p1", DailySessionStatus.ACTIVE),
            row("2026-10-30", "s2", "p1", DailySessionStatus.ACTIVE))
    result = summarize_patient_sessions(rows, start=date(2026, 10, 30), end=date(2026, 10, 30))
    assert len(result) == 1 and result[0].scheduled == 1
    with pytest.raises(ValueError, match="start date"):
        summarize_patient_sessions(rows, start=date(2026, 11, 1), end=date(2026, 10, 30))


def test_duplicate_session_rejected():
    rows = (row("2026-10-30", "S1", "p1", DailySessionStatus.ACTIVE),
            row("2026-10-30", "s1", "p1", DailySessionStatus.ACTIVE))
    with pytest.raises(ValueError, match="Duplicate"):
        summarize_patient_sessions(rows)


def test_unresolved_session_rejected():
    rows = (row("2026-10-30", "s1", "p1", DailySessionStatus.THERAPIST_ABSENT),)
    with pytest.raises(ValueError, match="Unresolved"):
        summarize_patient_sessions(rows)
