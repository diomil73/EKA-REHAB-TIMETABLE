from datetime import date, time

import pytest

from rehab_core.daily_state import build_daily_session_states
from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    DailySessionCancellation,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
    SessionCancellationKind,
)
from rehab_core.therapist_day_history import (
    TherapistDayHistoryError,
    build_therapist_day_history_rows,
)


DAY = date(2026, 10, 26)


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int, mm: int) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=time(hh, mm),
        treatment="ΦΘ",
    )


def _row(rows, therapist_id: str):
    return next(row for row in rows if row.therapist_id == therapist_id)


def test_shared_slot_counts_one_occupied_timeslot_but_two_active_sessions():
    states = build_daily_session_states(
        [
            _session("S1", "P1", "T1", 8, 30),
            _session("S2", "P2", "T1", 8, 30),
        ]
    )

    rows = build_therapist_day_history_rows(
        states,
        snapshot_id="snap-v1",
        service_date=DAY,
        therapist_ids=("T1",),
        capacity_timeslots_by_therapist={"T1": 6},
        approved_exception_events_by_therapist={"T1": 1},
    )
    row = _row(rows, "T1")

    assert row.planned_occupied_timeslots == 1
    assert row.planned_sessions == 2
    assert row.final_occupied_timeslots == 1
    assert row.final_active_sessions == 2
    assert row.capacity_timeslots == 6
    assert row.approved_shared_or_over_capacity_events == 1


def test_replacement_moves_final_workload_to_replacement_therapist():
    original = _session("S1", "P1", "T1", 9, 15)
    states = build_daily_session_states(
        [original],
        replacements=[
            ReplacementAssignment(
                replacement_id="R1",
                target_session_id="S1",
                patient_id="P1",
                original_therapist_id="T1",
                replacement_therapist_id="T2",
                replacement_date=DAY,
                replacement_time=time(10, 45),
            )
        ],
    )

    rows = build_therapist_day_history_rows(
        states,
        snapshot_id="snap-v1",
        service_date=DAY,
        therapist_ids=("T1", "T2"),
    )

    original_row = _row(rows, "T1")
    replacement_row = _row(rows, "T2")
    assert original_row.planned_sessions == 1
    assert original_row.final_active_sessions == 0
    assert original_row.final_occupied_timeslots == 0
    assert replacement_row.planned_sessions == 0
    assert replacement_row.final_active_sessions == 1
    assert replacement_row.final_occupied_timeslots == 1
    assert replacement_row.replacements_received == 1


def test_student_replacement_is_not_misreported_as_therapist_workload():
    original = _session("S1", "P1", "T1", 9, 15)
    states = build_daily_session_states(
        [original],
        replacements=[
            ReplacementAssignment(
                replacement_id="R1",
                target_session_id="S1",
                patient_id="P1",
                original_therapist_id="T1",
                replacement_therapist_id="STUDENT-1",
                replacement_date=DAY,
                replacement_time=time(9, 15),
                replacement_provider_kind=ReplacementProviderKind.STUDENT,
            )
        ],
    )

    rows = build_therapist_day_history_rows(
        states,
        snapshot_id="snap-v1",
        service_date=DAY,
        therapist_ids=("T1",),
    )

    assert {row.therapist_id for row in rows} == {"T1"}
    assert _row(rows, "T1").final_active_sessions == 0


def test_nonperformed_outcomes_are_attributed_to_original_therapist():
    patient_absent = _session("S-ABS", "P1", "T1", 8, 30)
    no_show = _session("S-NS", "P2", "T1", 9, 15)
    suspended = _session("S-SUS", "P3", "T1", 10, 0)
    states = build_daily_session_states(
        [patient_absent, no_show, suspended],
        absences=[
            DailyAbsence(
                absence_kind=AbsenceKind.PATIENT,
                subject_id="P1",
                absence_date=DAY,
            )
        ],
        cancellations=[
            DailySessionCancellation(
                cancellation_id="C-NS",
                target_session_id="S-NS",
                cancellation_date=DAY,
                kind=SessionCancellationKind.PATIENT_NO_SHOW,
            ),
            DailySessionCancellation(
                cancellation_id="C-SUS",
                target_session_id="S-SUS",
                cancellation_date=DAY,
                kind=SessionCancellationKind.TREATMENT_SUSPENDED,
            ),
        ],
    )

    row = _row(
        build_therapist_day_history_rows(
            states,
            snapshot_id="snap-v1",
            service_date=DAY,
        ),
        "T1",
    )

    assert row.planned_sessions == 3
    assert row.final_active_sessions == 0
    assert row.released_by_cancellation_or_no_show == 2
    assert row.treatment_suspensions_affecting_load == 1


def test_history_fails_closed_when_any_session_is_unresolved():
    session = _session("S-ORPH", "P1", "T1", 8, 30)
    states = build_daily_session_states(
        [session],
        absences=[
            DailyAbsence(
                absence_kind=AbsenceKind.THERAPIST,
                subject_id="T1",
                absence_date=DAY,
            )
        ],
    )

    with pytest.raises(TherapistDayHistoryError, match="S-ORPH"):
        build_therapist_day_history_rows(
            states,
            snapshot_id="snap-v1",
            service_date=DAY,
        )
