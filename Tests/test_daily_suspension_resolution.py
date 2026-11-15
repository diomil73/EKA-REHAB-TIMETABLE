from datetime import date, time

import pytest

from rehab_core.models import Patient, Session, SessionCancellationKind
from rehab_excel.daily_suspension_resolution import (
    DailySuspensionResolutionError,
    build_unresolved_suspensions,
)


TODAY = date(2026, 10, 26)


def _session(session_id: str, patient_id: str = "P1") -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id="T-ABSENT",
        session_date=TODAY,
        start_time=time(10, 0),
    )


def test_unresolved_solver_ids_become_explicit_treatment_suspensions():
    suspensions = build_unresolved_suspensions(
        ("S1",),
        sessions=(_session("S1"),),
        target_date=TODAY,
    )

    assert len(suspensions) == 1
    suspension = suspensions[0]
    assert suspension.target_session_id == "S1"
    assert suspension.cancellation_date == TODAY
    assert suspension.kind == SessionCancellationKind.TREATMENT_SUSPENDED
    assert "Αναστολή θεραπείας" in (suspension.reason or "")


def test_duplicate_unresolved_ids_do_not_create_duplicate_suspensions():
    suspensions = build_unresolved_suspensions(
        ("S1", "S1"),
        sessions=(_session("S1"),),
        target_date=TODAY,
    )

    assert [item.target_session_id for item in suspensions] == ["S1"]


def test_unknown_unresolved_session_is_a_safety_error():
    with pytest.raises(DailySuspensionResolutionError, match="does not exist"):
        build_unresolved_suspensions(
            ("MISSING",),
            sessions=(_session("S1"),),
            target_date=TODAY,
        )


def test_cross_date_unresolved_session_is_a_safety_error():
    other = Session(
        session_id="S1",
        patient_id="P1",
        therapist_id="T-ABSENT",
        session_date=date(2026, 10, 27),
        start_time=time(10, 0),
    )

    with pytest.raises(DailySuspensionResolutionError, match="not 2026-10-26"):
        build_unresolved_suspensions(
            ("S1",),
            sessions=(other,),
            target_date=TODAY,
        )
