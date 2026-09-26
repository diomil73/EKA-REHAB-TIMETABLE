from __future__ import annotations

from datetime import date, time
from typing import Iterable

from .models import (
    AbsenceKind,
    DailyAbsence,
    DailySessionCancellation,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
    StudentAssignment,
)


def _is_absent(
    *,
    absence_kind: AbsenceKind,
    subject_id: str,
    target_date: date,
    target_time: time,
    absences: Iterable[DailyAbsence],
) -> bool:
    return any(
        absence.absence_kind == absence_kind
        and absence.subject_id == subject_id
        and absence.covers(target_date, target_time)
        for absence in absences
    )


def _replacement_for_session(
    session_id: str,
    replacements: Iterable[ReplacementAssignment],
) -> ReplacementAssignment | None:
    for replacement in replacements:
        if replacement.target_session_id == session_id:
            return replacement
    return None


def _session_cancelled(
    session: Session,
    cancellations: Iterable[DailySessionCancellation],
) -> bool:
    return any(cancellation.applies_to(session) for cancellation in cancellations)


def is_therapist_available(
    therapist_id: str,
    target_date: date,
    target_time: time,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
) -> bool:
    """Return real operational availability at one exact date/time."""

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)
    cancellations = tuple(cancellations)

    if _is_absent(
        absence_kind=AbsenceKind.THERAPIST,
        subject_id=therapist_id,
        target_date=target_date,
        target_time=target_time,
        absences=absences,
    ):
        return False

    for replacement in replacements:
        if replacement.replacement_provider_kind != ReplacementProviderKind.THERAPIST:
            continue
        target_session = next(
            (session for session in sessions if session.session_id == replacement.target_session_id),
            None,
        )
        if target_session is not None and _session_cancelled(target_session, cancellations):
            continue
        if (
            replacement.replacement_therapist_id == therapist_id
            and replacement.replacement_date == target_date
            and replacement.replacement_time == target_time
        ):
            patient_absent = _is_absent(
                absence_kind=AbsenceKind.PATIENT,
                subject_id=replacement.patient_id,
                target_date=target_date,
                target_time=target_time,
                absences=absences,
            )
            if not patient_absent:
                return False

    for session in sessions:
        if (
            session.therapist_id != therapist_id
            or session.session_date != target_date
            or session.start_time != target_time
        ):
            continue

        if _session_cancelled(session, cancellations):
            continue
        if _replacement_for_session(session.session_id, replacements) is not None:
            continue

        patient_absent = _is_absent(
            absence_kind=AbsenceKind.PATIENT,
            subject_id=session.patient_id,
            target_date=target_date,
            target_time=target_time,
            absences=absences,
        )
        if not patient_absent:
            return False

    return True


def is_student_available(
    student_id: str,
    target_date: date,
    target_time: time,
    sessions: Iterable[Session],
    student_assignments: Iterable[StudentAssignment] = (),
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
) -> bool:
    """Return whether a student is free at one exact operational timeslot."""

    sessions = tuple(sessions)
    student_assignments = tuple(student_assignments)
    absences = tuple(absences)
    replacements = tuple(replacements)
    cancellations = tuple(cancellations)
    session_by_id = {session.session_id: session for session in sessions}

    if _is_absent(
        absence_kind=AbsenceKind.STUDENT,
        subject_id=student_id,
        target_date=target_date,
        target_time=target_time,
        absences=absences,
    ):
        return False

    for replacement in replacements:
        if replacement.replacement_provider_kind != ReplacementProviderKind.STUDENT:
            continue
        target_session = session_by_id.get(replacement.target_session_id)
        if target_session is not None and _session_cancelled(target_session, cancellations):
            continue
        if (
            replacement.replacement_therapist_id == student_id
            and replacement.replacement_date == target_date
            and replacement.replacement_time == target_time
        ):
            if not _is_absent(
                absence_kind=AbsenceKind.PATIENT,
                subject_id=replacement.patient_id,
                target_date=target_date,
                target_time=target_time,
                absences=absences,
            ):
                return False

    for assignment in student_assignments:
        if assignment.student_id != student_id:
            continue
        session = session_by_id.get(assignment.session_id)
        if session is None:
            continue
        if session.session_date != target_date or session.start_time != target_time:
            continue
        if _session_cancelled(session, cancellations):
            continue
        if _is_absent(
            absence_kind=AbsenceKind.PATIENT,
            subject_id=session.patient_id,
            target_date=target_date,
            target_time=target_time,
            absences=absences,
        ):
            continue
        return False

    return True


def is_patient_available(
    patient_id: str,
    target_date: date,
    target_time: time,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    *,
    ignore_session_id: str | None = None,
) -> bool:
    """Return whether a patient can receive a session at a date/time."""

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)
    cancellations = tuple(cancellations)
    session_by_id = {session.session_id: session for session in sessions}

    if _is_absent(
        absence_kind=AbsenceKind.PATIENT,
        subject_id=patient_id,
        target_date=target_date,
        target_time=target_time,
        absences=absences,
    ):
        return False

    for replacement in replacements:
        if replacement.target_session_id == ignore_session_id:
            continue
        target_session = session_by_id.get(replacement.target_session_id)
        if target_session is not None and _session_cancelled(target_session, cancellations):
            continue
        if (
            replacement.patient_id == patient_id
            and replacement.replacement_date == target_date
            and replacement.replacement_time == target_time
        ):
            return False

    for session in sessions:
        if session.session_id == ignore_session_id:
            continue
        if (
            session.patient_id != patient_id
            or session.session_date != target_date
            or session.start_time != target_time
        ):
            continue

        if _session_cancelled(session, cancellations):
            continue
        if _replacement_for_session(session.session_id, replacements) is not None:
            continue

        return False

    return True
