from datetime import date, time

from rehab_core.daily_state import build_daily_session_states
from rehab_core.daily_statistics import calculate_daily_treatment_statistics
from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    DailySessionCancellation,
    ReplacementAssignment,
    Session,
    SessionCancellationKind,
)
from rehab_core.orphan_resolution import create_treatment_suspension


DAY = date(2026, 10, 19)


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int, mm: int) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=time(hh, mm),
        treatment="ΦΘ",
    )


def test_statistics_distinguish_performed_cancel_reasons_and_unresolved_orphans():
    active = _session("S-ACT", "P-ACT", "T1", 8, 30)
    replaced = _session("S-REP", "P-REP", "T2", 9, 15)
    no_show = _session("S-NS", "P-NS", "T3", 10, 0)
    suspended = _session("S-SUS", "P-SUS", "T4", 10, 45)
    unresolved = _session("S-ORPH", "P-ORPH", "T5", 11, 30)

    states = build_daily_session_states(
        [active, replaced, no_show, suspended, unresolved],
        replacements=[
            ReplacementAssignment(
                replacement_id="R1",
                target_session_id=replaced.session_id,
                patient_id=replaced.patient_id,
                original_therapist_id=replaced.therapist_id,
                replacement_therapist_id="T9",
                replacement_date=DAY,
                replacement_time=time(12, 15),
            )
        ],
        cancellations=[
            DailySessionCancellation(
                cancellation_id="C-NS",
                target_session_id=no_show.session_id,
                cancellation_date=DAY,
                kind=SessionCancellationKind.PATIENT_NO_SHOW,
                reason="Δεν προσήλθε",
            ),
            create_treatment_suspension(suspended),
        ],
        absences=[
            DailyAbsence(
                absence_kind=AbsenceKind.THERAPIST,
                subject_id=unresolved.therapist_id,
                absence_date=DAY,
                reason="Απουσία θεραπευτή",
            )
        ],
    )

    stats = calculate_daily_treatment_statistics(states)

    assert stats.scheduled_sessions == 5
    assert stats.performed_sessions == 2
    assert stats.replaced_sessions == 1
    assert stats.cancelled_sessions == 2
    assert stats.patient_no_show_sessions == 1
    assert stats.treatment_suspended_sessions == 1
    assert stats.department_postponed_sessions == 0
    assert stats.patient_absent_sessions == 0
    assert stats.non_performed_sessions == 2
    assert stats.unresolved_sessions == 1
    assert stats.accounted_sessions == 5
    assert stats.is_finalized is False


def test_statistics_finalize_once_last_orphan_is_explicitly_suspended():
    session = _session("S1", "P1", "T1", 9, 15)
    states = build_daily_session_states(
        [session],
        cancellations=[create_treatment_suspension(session)],
    )

    stats = calculate_daily_treatment_statistics(states)

    assert stats.scheduled_sessions == 1
    assert stats.performed_sessions == 0
    assert stats.treatment_suspended_sessions == 1
    assert stats.unresolved_sessions == 0
    assert stats.is_finalized is True
