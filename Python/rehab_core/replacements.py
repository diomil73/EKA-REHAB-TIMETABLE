from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable

from .availability import is_therapist_available
from .models import (
    AbsenceKind,
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    Session,
    Therapist,
)


@dataclass(frozen=True)
class ReplacementCandidate:
    therapist_id: str
    display_name: str
    active_sessions: int
    infectious_sessions: int


def _patient_absent(
    patient_id: str,
    session: Session,
    absences: Iterable[DailyAbsence],
) -> bool:
    return any(
        absence.absence_kind == AbsenceKind.PATIENT
        and absence.subject_id == patient_id
        and absence.covers(session.session_date, session.start_time)
        for absence in absences
    )


def _daily_workload(
    therapist_id: str,
    target_date: date,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence],
    replacements: Iterable[ReplacementAssignment],
    patient_by_id: dict[str, Patient],
) -> tuple[int, int]:
    """Return operational (active sessions, infectious sessions) for the day."""

    active_sessions = 0
    infectious_sessions = 0

    for session in sessions:
        if session.therapist_id != therapist_id or session.session_date != target_date:
            continue
        if _patient_absent(session.patient_id, session, absences):
            continue

        active_sessions += 1
        patient = patient_by_id.get(session.patient_id)
        if patient is not None and patient.infectious:
            infectious_sessions += 1

    for replacement in replacements:
        if (
            replacement.replacement_therapist_id != therapist_id
            or replacement.replacement_date != target_date
        ):
            continue

        active_sessions += 1
        patient = patient_by_id.get(replacement.patient_id)
        if patient is not None and patient.infectious:
            infectious_sessions += 1

    return active_sessions, infectious_sessions


def find_replacement_candidates(
    target_session: Session,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> list[ReplacementCandidate]:
    """Return available replacement therapists in deterministic order.

    The ranking uses today's operational state, including already accepted
    replacement assignments. The base schedule remains unchanged.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    patients = tuple(patients)
    replacements = tuple(replacements)
    patient_by_id = {patient.patient_id: patient for patient in patients}
    target_patient = patient_by_id.get(target_session.patient_id)
    target_is_infectious = bool(target_patient and target_patient.infectious)

    candidates: list[ReplacementCandidate] = []

    for therapist in therapists:
        if therapist.therapist_id == target_session.therapist_id:
            continue
        if target_session.robotic and not therapist.robotic_capable:
            continue
        if not is_therapist_available(
            therapist_id=therapist.therapist_id,
            target_date=target_session.session_date,
            target_time=target_session.start_time,
            sessions=sessions,
            absences=absences,
            replacements=replacements,
        ):
            continue

        active_sessions, infectious_sessions = _daily_workload(
            therapist_id=therapist.therapist_id,
            target_date=target_session.session_date,
            sessions=sessions,
            absences=absences,
            replacements=replacements,
            patient_by_id=patient_by_id,
        )
        candidates.append(
            ReplacementCandidate(
                therapist_id=therapist.therapist_id,
                display_name=therapist.display_name,
                active_sessions=active_sessions,
                infectious_sessions=infectious_sessions,
            )
        )

    if target_is_infectious:
        candidates.sort(
            key=lambda candidate: (
                candidate.infectious_sessions,
                candidate.active_sessions,
                candidate.display_name.casefold(),
                candidate.therapist_id,
            )
        )
    else:
        candidates.sort(
            key=lambda candidate: (
                candidate.active_sessions,
                candidate.infectious_sessions,
                candidate.display_name.casefold(),
                candidate.therapist_id,
            )
        )

    return candidates


def create_replacement_assignment(
    *,
    replacement_id: str,
    target_session: Session,
    replacement_therapist_id: str,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    reason: str | None = None,
) -> ReplacementAssignment:
    """Validate and create a daily replacement overlay.

    This function does not edit the base Session. It raises ValueError when
    the requested therapist cannot legally take the session in today's state.
    """

    therapist_by_id = {
        therapist.therapist_id: therapist for therapist in therapists
    }
    therapist = therapist_by_id.get(replacement_therapist_id)
    if therapist is None:
        raise ValueError("Unknown replacement therapist")
    if replacement_therapist_id == target_session.therapist_id:
        raise ValueError("Replacement therapist cannot be the original therapist")
    if target_session.robotic and not therapist.robotic_capable:
        raise ValueError("Replacement therapist is not robotic-capable")

    if not is_therapist_available(
        therapist_id=replacement_therapist_id,
        target_date=target_session.session_date,
        target_time=target_session.start_time,
        sessions=sessions,
        absences=absences,
        replacements=replacements,
    ):
        raise ValueError("Replacement therapist is not available")

    return ReplacementAssignment(
        replacement_id=replacement_id,
        target_session_id=target_session.session_id,
        patient_id=target_session.patient_id,
        original_therapist_id=target_session.therapist_id,
        replacement_therapist_id=replacement_therapist_id,
        replacement_date=target_session.session_date,
        replacement_time=target_session.start_time,
        reason=reason,
    )
