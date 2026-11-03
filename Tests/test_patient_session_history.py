from datetime import date, time

import pytest

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import ReplacementProviderKind, SessionCancellationKind
from rehab_core.patient_session_history import (
    PatientSessionHistoryError,
    build_patient_session_history_rows,
)


SERVICE_DATE = date(2026, 10, 30)


def _state(
    session_id: str,
    patient_id: str,
    status: DailySessionStatus,
    *,
    original_time: time,
    original_therapist: str = "T1",
    effective_therapist: str | None = "T1",
    effective_time: time | None = None,
    replacement_id: str | None = None,
    cancellation_id: str | None = None,
    cancellation_kind: SessionCancellationKind | None = None,
    reason: str | None = None,
) -> DailySessionState:
    if effective_time is None and status in {DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED}:
        effective_time = original_time
    return DailySessionState(
        session_id=session_id,
        patient_id=patient_id,
        session_date=SERVICE_DATE,
        status=status,
        original_therapist_id=original_therapist,
        original_time=original_time,
        effective_therapist_id=effective_therapist,
        effective_time=effective_time,
        replacement_id=replacement_id,
        reason=reason,
        effective_provider_kind=(
            ReplacementProviderKind.THERAPIST
            if effective_therapist is not None
            else None
        ),
        cancellation_id=cancellation_id,
        cancellation_kind=cancellation_kind,
    )


def test_freezes_original_and_effective_session_outcomes_deterministically():
    states = (
        _state(
            "S2",
            "P2",
            DailySessionStatus.REPLACED,
            original_time=time(10, 0),
            original_therapist="T1",
            effective_therapist="T2",
            effective_time=time(10, 30),
            replacement_id="R1",
            reason="therapist absence",
        ),
        _state(
            "S1",
            "P1",
            DailySessionStatus.ACTIVE,
            original_time=time(9, 0),
        ),
    )

    rows = build_patient_session_history_rows(
        states,
        snapshot_id="finalized:2026-10-30:v2",
        service_date=SERVICE_DATE,
        snapshot_version=2,
    )

    assert [row.session_id for row in rows] == ["S1", "S2"]
    assert rows[0].patient_id == "P1"
    assert rows[0].status == DailySessionStatus.ACTIVE
    assert rows[1].snapshot_version == 2
    assert rows[1].original_therapist_id == "T1"
    assert rows[1].original_time == time(10, 0)
    assert rows[1].effective_therapist_id == "T2"
    assert rows[1].effective_time == time(10, 30)
    assert rows[1].replacement_id == "R1"


def test_preserves_patient_absence_and_cancellation_reason_details():
    states = (
        _state(
            "S1",
            "P1",
            DailySessionStatus.PATIENT_ABSENT,
            original_time=time(9, 0),
            effective_therapist=None,
            reason="hospitalized",
        ),
        _state(
            "S2",
            "P2",
            DailySessionStatus.CANCELLED,
            original_time=time(10, 0),
            effective_therapist=None,
            cancellation_id="C1",
            cancellation_kind=SessionCancellationKind.PATIENT_NO_SHOW,
            reason="no show",
        ),
    )

    rows = build_patient_session_history_rows(
        states,
        snapshot_id="finalized:2026-10-30:v1",
        service_date=SERVICE_DATE,
        snapshot_version=1,
    )

    assert rows[0].reason == "hospitalized"
    assert rows[0].cancellation_kind is None
    assert rows[1].cancellation_id == "C1"
    assert rows[1].cancellation_kind == SessionCancellationKind.PATIENT_NO_SHOW.value
    assert rows[1].reason == "no show"


def test_refuses_unresolved_therapist_absence():
    state = _state(
        "S1",
        "P1",
        DailySessionStatus.THERAPIST_ABSENT,
        original_time=time(9, 0),
        effective_therapist=None,
        reason="sick leave",
    )

    with pytest.raises(PatientSessionHistoryError, match="unresolved sessions"):
        build_patient_session_history_rows(
            (state,),
            snapshot_id="finalized:2026-10-30:v1",
            service_date=SERVICE_DATE,
            snapshot_version=1,
        )


def test_refuses_wrong_service_date_and_duplicate_sessions():
    wrong_date = DailySessionState(
        session_id="S1",
        patient_id="P1",
        session_date=date(2026, 10, 29),
        status=DailySessionStatus.ACTIVE,
        original_therapist_id="T1",
        original_time=time(9, 0),
        effective_therapist_id="T1",
        effective_time=time(9, 0),
        effective_provider_kind=ReplacementProviderKind.THERAPIST,
    )
    with pytest.raises(PatientSessionHistoryError, match="another service date"):
        build_patient_session_history_rows(
            (wrong_date,),
            snapshot_id="finalized:2026-10-30:v1",
            service_date=SERVICE_DATE,
            snapshot_version=1,
        )

    duplicate_a = _state("S1", "P1", DailySessionStatus.ACTIVE, original_time=time(9, 0))
    duplicate_b = _state("s1", "P2", DailySessionStatus.ACTIVE, original_time=time(10, 0))
    with pytest.raises(PatientSessionHistoryError, match="Duplicate session_id"):
        build_patient_session_history_rows(
            (duplicate_a, duplicate_b),
            snapshot_id="finalized:2026-10-30:v1",
            service_date=SERVICE_DATE,
            snapshot_version=1,
        )


def test_refuses_invalid_snapshot_identity():
    state = _state("S1", "P1", DailySessionStatus.ACTIVE, original_time=time(9, 0))
    with pytest.raises(PatientSessionHistoryError, match="snapshot_id"):
        build_patient_session_history_rows(
            (state,), snapshot_id=" ", service_date=SERVICE_DATE, snapshot_version=1
        )
    with pytest.raises(PatientSessionHistoryError, match="snapshot_version"):
        build_patient_session_history_rows(
            (state,), snapshot_id="snap", service_date=SERVICE_DATE, snapshot_version=0
        )
