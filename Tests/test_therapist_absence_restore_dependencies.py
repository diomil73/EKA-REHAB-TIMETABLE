from datetime import date, time

from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    DailySessionCancellation,
    ReplacementProviderKind,
    Session,
    SessionCancellationKind,
)
from rehab_excel.daily_replacement_facts import DailyReplacementFactRow
from rehab_excel.daily_suspension_resolution import SUSPENSION_REASON
from rehab_excel.therapist_absence_restore import (
    dependent_active_replacement_ids,
    dependent_generated_suspension_ids,
    released_therapist_session_ids,
)


DAY = date(2026, 10, 27)


def _session(session_id: str, therapist: str, hh: int) -> Session:
    return Session(
        session_id=session_id,
        patient_id=f"P-{session_id}",
        therapist_id=therapist,
        session_date=DAY,
        start_time=time(hh, 0),
    )


def _absence(therapist: str, start: int | None = None, end: int | None = None) -> DailyAbsence:
    return DailyAbsence(
        absence_kind=AbsenceKind.THERAPIST,
        subject_id=therapist,
        absence_date=DAY,
        start_time=None if start is None else time(start, 0),
        end_time=None if end is None else time(end, 0),
        reason="test",
    )


def _fact(session: Session, replacement_id: str, status: str = "ACTIVE") -> DailyReplacementFactRow:
    return DailyReplacementFactRow(
        replacement_id=replacement_id,
        target_session_id=session.session_id,
        patient_id=session.patient_id,
        original_therapist_id=session.therapist_id,
        replacement_therapist_id="OTHER",
        replacement_date=DAY,
        replacement_time=time(12, 0),
        reason="test",
        provider_kind=ReplacementProviderKind.THERAPIST,
        status=status,
    )


def test_restore_releases_only_sessions_no_longer_covered_by_any_therapist_absence():
    s9 = _session("S9", "GAVRAS", 9)
    s10 = _session("S10", "GAVRAS", 10)
    s11 = _session("S11", "GAVRAS", 11)

    selected_row = _absence("GAVRAS", 9, 11)
    overlapping_row = _absence("GAVRAS", 10, 12)

    released = released_therapist_session_ids(
        sessions=(s9, s10, s11),
        before_absences=(selected_row, overlapping_row),
        after_absences=(overlapping_row,),
    )

    assert released == ("S9",)


def test_all_day_restore_releases_only_original_therapist_sessions():
    gavras = _session("G1", "GAVRAS", 9)
    gounelas = _session("G2", "GOUNELAS", 9)

    released = released_therapist_session_ids(
        sessions=(gavras, gounelas),
        before_absences=(_absence("GAVRAS"),),
        after_absences=(),
    )

    assert released == ("G1",)


def test_only_active_replacements_for_released_sessions_are_restored():
    released = _session("S1", "GAVRAS", 9)
    unrelated = _session("S2", "GOUNELAS", 10)

    ids = dependent_active_replacement_ids(
        (
            _fact(released, "R1", "ACTIVE"),
            _fact(released, "R1-old", "RESTORED"),
            _fact(unrelated, "R2", "ACTIVE"),
        ),
        released_session_ids=("S1",),
    )

    assert ids == ("R1",)


def test_only_solver_generated_suspensions_for_released_sessions_are_removed():
    generated = DailySessionCancellation(
        cancellation_id="global-suspension:S1:2026-10-27",
        target_session_id="S1",
        cancellation_date=DAY,
        kind=SessionCancellationKind.TREATMENT_SUSPENDED,
        reason=f"Αναστολή θεραπείας | {SUSPENSION_REASON}",
    )
    manual = DailySessionCancellation(
        cancellation_id="daily-input:manual",
        target_session_id="S1",
        cancellation_date=DAY,
        kind=SessionCancellationKind.TREATMENT_SUSPENDED,
        reason="Χειροκίνητη αναστολή",
    )
    other = DailySessionCancellation(
        cancellation_id="global-suspension:S2:2026-10-27",
        target_session_id="S2",
        cancellation_date=DAY,
        kind=SessionCancellationKind.TREATMENT_SUSPENDED,
        reason=f"Αναστολή θεραπείας | {SUSPENSION_REASON}",
    )

    ids = dependent_generated_suspension_ids(
        (generated, manual, other),
        released_session_ids=("S1",),
    )

    assert ids == ("global-suspension:S1:2026-10-27",)
