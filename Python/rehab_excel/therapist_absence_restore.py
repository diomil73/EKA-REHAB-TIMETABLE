from __future__ import annotations

from collections.abc import Iterable
from datetime import time

from rehab_core.models import AbsenceKind, DailyAbsence, DailySessionCancellation, Session, SessionCancellationKind
from rehab_excel.daily_replacement_facts import DailyReplacementFactRow
from rehab_excel.daily_suspension_resolution import SUSPENSION_REASON


def _covers(absence: DailyAbsence, session: Session) -> bool:
    if absence.absence_kind != AbsenceKind.THERAPIST:
        return False
    if absence.subject_id != session.therapist_id:
        return False
    if absence.absence_date != session.session_date:
        return False
    if absence.start_time is None and absence.end_time is None:
        return True
    start = absence.start_time or time.min
    end = absence.end_time or time.max
    return start <= session.start_time < end


def released_therapist_session_ids(
    *,
    sessions: Iterable[Session],
    before_absences: Iterable[DailyAbsence],
    after_absences: Iterable[DailyAbsence],
) -> tuple[str, ...]:
    """Return sessions released specifically by removing one therapist-absence source fact.

    A session is released only when it was covered by a therapist absence before the
    restore and is no longer covered by any therapist absence afterwards. This makes
    restore safe when absence rows overlap or another therapist-absence fact still
    applies to the same session.
    """

    before = tuple(before_absences)
    after = tuple(after_absences)
    released: list[str] = []
    for session in sessions:
        was_absent = any(_covers(absence, session) for absence in before)
        remains_absent = any(_covers(absence, session) for absence in after)
        if was_absent and not remains_absent:
            released.append(session.session_id)
    return tuple(released)


def dependent_active_replacement_ids(
    replacement_facts: Iterable[DailyReplacementFactRow],
    *,
    released_session_ids: Iterable[str],
) -> tuple[str, ...]:
    released = {str(value).strip() for value in released_session_ids}
    return tuple(
        fact.replacement_id
        for fact in replacement_facts
        if fact.is_active and fact.target_session_id in released
    )


def dependent_generated_suspension_ids(
    cancellations: Iterable[DailySessionCancellation],
    *,
    released_session_ids: Iterable[str],
) -> tuple[str, ...]:
    """Return only solver-generated treatment suspensions for released sessions."""

    released = {str(value).strip() for value in released_session_ids}
    return tuple(
        item.cancellation_id
        for item in cancellations
        if item.target_session_id in released
        and item.kind == SessionCancellationKind.TREATMENT_SUSPENDED
        and SUSPENSION_REASON in str(item.reason or "")
    )
