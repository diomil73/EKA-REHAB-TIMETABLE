from __future__ import annotations

from datetime import date, time

import pytest

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.history import (
    DayFinalizationError,
    build_finalized_day_snapshot,
)
from rehab_core.models import SessionCancellationKind


TARGET_DATE = date(2026, 10, 29)


def _state(
    session_id: str,
    *,
    patient_id: str,
    status: DailySessionStatus,
    original_therapist_id: str,
    original_time: time,
    effective_therapist_id: str | None,
    effective_time: time | None,
    replacement_id: str | None = None,
    cancellation_id: str | None = None,
    cancellation_kind: SessionCancellationKind | None = None,
) -> DailySessionState:
    return DailySessionState(
        session_id=session_id,
        patient_id=patient_id,
        session_date=TARGET_DATE,
        status=status,
        original_therapist_id=original_therapist_id,
        original_time=original_time,
        effective_therapist_id=effective_therapist_id,
        effective_time=effective_time,
        replacement_id=replacement_id,
        cancellation_id=cancellation_id,
        cancellation_kind=cancellation_kind,
    )


def test_finalized_snapshot_preserves_original_and_effective_replacement_truth():
    states = [
        _state(
            "S1",
            patient_id="P1",
            status=DailySessionStatus.ACTIVE,
            original_therapist_id="T1",
            original_time=time(9, 0),
            effective_therapist_id="T1",
            effective_time=time(9, 0),
        ),
        _state(
            "S2",
            patient_id="P2",
            status=DailySessionStatus.REPLACED,
            original_therapist_id="T1",
            original_time=time(10, 0),
            effective_therapist_id="T2",
            effective_time=time(12, 15),
            replacement_id="R1",
        ),
    ]

    snapshot = build_finalized_day_snapshot(TARGET_DATE, states)

    assert snapshot.version == 1
    assert snapshot.treatment_statistics.scheduled_sessions == 2
    assert snapshot.treatment_statistics.performed_sessions == 2
    assert snapshot.treatment_statistics.replaced_sessions == 1
    assert snapshot.treatment_statistics.is_finalized is True

    replacement = next(item for item in snapshot.sessions if item.session_id == "S2")
    assert replacement.original_therapist_id == "T1"
    assert replacement.original_time == time(10, 0)
    assert replacement.effective_therapist_id == "T2"
    assert replacement.effective_time == time(12, 15)
    assert replacement.replacement_id == "R1"

    therapist = {row.therapist_id: row for row in snapshot.therapist_history}
    assert therapist["T1"].planned_session_count == 2
    assert therapist["T1"].actual_session_count == 1
    assert therapist["T1"].occupied_timeslots == (time(9, 0),)
    assert therapist["T2"].planned_session_count == 0
    assert therapist["T2"].actual_session_count == 1
    assert therapist["T2"].replacement_session_count == 1
    assert therapist["T2"].occupied_timeslots == (time(12, 15),)


def test_close_day_rejects_unresolved_therapist_absence():
    states = [
        _state(
            "S1",
            patient_id="P1",
            status=DailySessionStatus.THERAPIST_ABSENT,
            original_therapist_id="T1",
            original_time=time(10, 0),
            effective_therapist_id=None,
            effective_time=None,
        )
    ]

    with pytest.raises(DayFinalizationError) as exc_info:
        build_finalized_day_snapshot(TARGET_DATE, states)

    assert exc_info.value.unresolved_session_ids == ("S1",)


def test_close_day_rejects_orphan_sessions_outside_normal_daily_state():
    states = [
        _state(
            "S1",
            patient_id="P1",
            status=DailySessionStatus.ACTIVE,
            original_therapist_id="T1",
            original_time=time(9, 0),
            effective_therapist_id="T1",
            effective_time=time(9, 0),
        )
    ]

    with pytest.raises(DayFinalizationError) as exc_info:
        build_finalized_day_snapshot(
            TARGET_DATE,
            states,
            orphan_session_ids=("ORPHAN-1", "ORPHAN-1", "ORPHAN-2"),
        )

    assert exc_info.value.orphan_session_ids == ("ORPHAN-1", "ORPHAN-2")


def test_close_day_rejects_states_from_another_date():
    state = DailySessionState(
        session_id="S1",
        patient_id="P1",
        session_date=date(2026, 10, 28),
        status=DailySessionStatus.ACTIVE,
        original_therapist_id="T1",
        original_time=time(9, 0),
        effective_therapist_id="T1",
        effective_time=time(9, 0),
    )

    with pytest.raises(DayFinalizationError) as exc_info:
        build_finalized_day_snapshot(TARGET_DATE, [state])

    assert exc_info.value.wrong_date_session_ids == ("S1",)


def test_cancellation_outcomes_are_preserved_and_do_not_count_as_actual_workload():
    states = [
        _state(
            "S1",
            patient_id="P1",
            status=DailySessionStatus.CANCELLED,
            original_therapist_id="T1",
            original_time=time(11, 0),
            effective_therapist_id=None,
            effective_time=None,
            cancellation_id="C1",
            cancellation_kind=SessionCancellationKind.TREATMENT_SUSPENDED,
        )
    ]

    snapshot = build_finalized_day_snapshot(TARGET_DATE, states, version=3)

    assert snapshot.version == 3
    assert snapshot.treatment_statistics.treatment_suspended_sessions == 1
    fact = snapshot.sessions[0]
    assert fact.cancellation_id == "C1"
    assert fact.cancellation_kind == SessionCancellationKind.TREATMENT_SUSPENDED.value

    therapist = snapshot.therapist_history[0]
    assert therapist.planned_session_count == 1
    assert therapist.actual_session_count == 0
    assert therapist.cancelled_session_count == 1
    assert therapist.occupied_timeslots == ()


def test_snapshot_order_is_deterministic():
    states = [
        _state(
            "S-B",
            patient_id="P2",
            status=DailySessionStatus.ACTIVE,
            original_therapist_id="T1",
            original_time=time(10, 0),
            effective_therapist_id="T1",
            effective_time=time(10, 0),
        ),
        _state(
            "S-A",
            patient_id="P1",
            status=DailySessionStatus.ACTIVE,
            original_therapist_id="T2",
            original_time=time(9, 0),
            effective_therapist_id="T2",
            effective_time=time(9, 0),
        ),
    ]

    first = build_finalized_day_snapshot(TARGET_DATE, states)
    second = build_finalized_day_snapshot(TARGET_DATE, reversed(states))

    assert first == second
    assert [item.session_id for item in first.sessions] == ["S-A", "S-B"]


def test_snapshot_version_must_be_positive():
    with pytest.raises(ValueError, match="version"):
        build_finalized_day_snapshot(TARGET_DATE, [], version=0)
