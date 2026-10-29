from datetime import date, time

import pytest

from rehab_core.models import Session
from tools.prepare_orphan_smoke_input import choose_smoke_therapist


DAY = date(2026, 10, 19)


def _session(session_id: str, patient_id: str, therapist_id: str, at: time) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=at,
        treatment="ΦΘ",
    )


def test_choose_smoke_therapist_prefers_highest_session_count_and_sorts_affected():
    sessions = [
        _session("S3", "P3", "T2", time(10, 0)),
        _session("S1", "P1", "T1", time(9, 15)),
        _session("S4", "P4", "T2", time(8, 30)),
        _session("S5", "P5", "T2", time(9, 15)),
        _session("S2", "P2", "T1", time(10, 0)),
    ]

    therapist_id, affected = choose_smoke_therapist(
        sessions,
        allowed_therapists=("T1", "T2"),
        min_sessions=2,
    )

    assert therapist_id == "T2"
    assert [session.session_id for session in affected] == ["S4", "S5", "S3"]


def test_choose_smoke_therapist_ignores_non_allowed_providers():
    sessions = [
        _session("S1", "P1", "STUDENT", time(8, 30)),
        _session("S2", "P2", "STUDENT", time(9, 15)),
        _session("S3", "P3", "T1", time(10, 0)),
        _session("S4", "P4", "T1", time(10, 45)),
    ]

    therapist_id, affected = choose_smoke_therapist(
        sessions,
        allowed_therapists=("T1",),
        min_sessions=2,
    )

    assert therapist_id == "T1"
    assert len(affected) == 2


def test_choose_smoke_therapist_requires_multi_orphan_case():
    with pytest.raises(ValueError, match="at least 2"):
        choose_smoke_therapist([], allowed_therapists=("T1",), min_sessions=1)


def test_choose_smoke_therapist_stops_when_no_provider_has_enough_sessions():
    sessions = [_session("S1", "P1", "T1", time(8, 30))]

    with pytest.raises(ValueError, match="No therapist has at least 2"):
        choose_smoke_therapist(
            sessions,
            allowed_therapists=("T1",),
            min_sessions=2,
        )
