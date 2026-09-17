from __future__ import annotations

from datetime import date, time
from typing import Iterable

from .models import AbsenceKind, DailyAbsence, Session


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


def is_therapist_available(
    therapist_id: str,
    target_date: date,
    target_time: time,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
) -> bool:
    """Return real availability for the requested day/time.

    Rules implemented in milestone 1:
    - An absent therapist is unavailable.
    - A scheduled session blocks the therapist unless that session's patient
      is absent at the same date/time.
    - Base schedule data is not mutated.
    """

    if _is_absent(
        absence_kind=AbsenceKind.THERAPIST,
        subject_id=therapist_id,
        target_date=target_date,
        target_time=target_time,
        absences=absences,
    ):
        return False

    for session in sessions:
        if (
            session.therapist_id != therapist_id
            or session.session_date != target_date
            or session.start_time != target_time
        ):
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
