from __future__ import annotations

from datetime import date, time
from typing import Iterable

from .models import AbsenceKind, DailyAbsence, ReplacementAssignment, Session


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


def is_therapist_available(
    therapist_id: str,
    target_date: date,
    target_time: time,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> bool:
    """Return real operational availability for the requested day/time.

    Rules:
    - an absent therapist is unavailable;
    - an accepted replacement assignment blocks the replacement therapist;
    - once a base session has a replacement overlay, that base session no
      longer blocks its original therapist;
    - a normal base session blocks the therapist unless that session's patient
      is absent at the same date/time;
    - the base schedule is never mutated by daily exceptions.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)

    if _is_absent(
        absence_kind=AbsenceKind.THERAPIST,
        subject_id=therapist_id,
        target_date=target_date,
        target_time=target_time,
        absences=absences,
    ):
        return False

    for replacement in replacements:
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

        # The daily overlay replaces the base occurrence operationally. The
        # original therapist is therefore not considered occupied by it.
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


def is_patient_available(
    patient_id: str,
    target_date: date,
    target_time: time,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    *,
    ignore_session_id: str | None = None,
) -> bool:
    """Return whether a patient can receive a session at a date/time.

    Daily overlays are treated as the operational truth while the base
    schedule remains untouched. ``ignore_session_id`` is used when validating
    a replacement for the session being moved/reassigned, so the session does
    not conflict with itself.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)

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

        # If this base occurrence was replaced, its old base position is no
        # longer operationally occupied by the patient.
        if _replacement_for_session(session.session_id, replacements) is not None:
            continue

        return False

    return True
