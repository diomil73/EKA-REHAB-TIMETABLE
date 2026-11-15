from datetime import date, time

from rehab_core.daily_state import DailySessionStatus, build_daily_session_states
from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    Patient,
    Session,
    SessionCancellationKind,
    Therapist,
)
from rehab_core.orphan_resolution import (
    TREATMENT_SUSPENSION_REASON,
    create_treatment_suspension,
)
from rehab_core.therapist_absence_queue import (
    build_therapist_absence_replacement_queue,
)


def _session() -> Session:
    return Session(
        session_id="S1",
        patient_id="P1",
        therapist_id="T1",
        session_date=date(2026, 10, 19),
        start_time=time(9, 15),
        treatment="ΦΘ",
    )


def test_treatment_suspension_is_date_specific_cancelled_state_and_releases_slot():
    session = _session()
    suspension = create_treatment_suspension(session)

    assert suspension.kind == SessionCancellationKind.TREATMENT_SUSPENDED
    assert suspension.reason == TREATMENT_SUSPENSION_REASON
    assert suspension.cancellation_date == session.session_date
    assert suspension.target_session_id == session.session_id

    state = build_daily_session_states(
        [session],
        cancellations=[suspension],
    )[0]

    assert state.status == DailySessionStatus.CANCELLED
    assert state.cancellation_kind == SessionCancellationKind.TREATMENT_SUSPENDED
    assert state.reason == "Αναστολή θεραπείας"
    assert state.effective_therapist_id is None
    assert state.effective_time is None
    assert state.releases_provider_slot is True


def test_suspended_orphan_is_removed_from_replacement_queue_as_explicit_cancelled_outcome():
    session = _session()
    absence = DailyAbsence(
        absence_kind=AbsenceKind.THERAPIST,
        subject_id="T1",
        absence_date=session.session_date,
        reason="Απουσία θεραπευτή",
    )
    suspension = create_treatment_suspension(session)

    queue = build_therapist_absence_replacement_queue(
        sessions=[session],
        absences=[absence],
        therapists=[Therapist("T1", "Θεραπευτής 1")],
        patients=[Patient("P1", "Ασθενής 1")],
        timeslots=[time(9, 15), time(10, 0)],
        cancellations=[suspension],
    )

    assert queue.affected_sessions == 0
    assert queue.skipped_cancelled == 1
