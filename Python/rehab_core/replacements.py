from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Iterable

from .availability import is_patient_available, is_therapist_available
from .models import (
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    Session,
    Therapist,
)
from .workload import calculate_therapist_workload


@dataclass(frozen=True)
class ReplacementCandidate:
    therapist_id: str
    display_name: str
    active_sessions: int
    infectious_sessions: int
    robotic_sessions: int
    replacement_sessions: int


def find_replacement_candidates(
    target_session: Session,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    *,
    replacement_time: time | None = None,
) -> list[ReplacementCandidate]:
    """Return available replacement therapists in deterministic order.

    The ranking uses today's operational state, including already accepted
    replacement assignments. The base schedule remains unchanged.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    patients = tuple(patients)
    replacements = tuple(replacements)
    target_time = replacement_time or target_session.start_time
    patient_by_id = {patient.patient_id: patient for patient in patients}
    target_patient = patient_by_id.get(target_session.patient_id)
    target_is_infectious = bool(target_patient and target_patient.infectious)

    # There is no meaningful replacement if the patient cannot attend at the
    # proposed operational time.
    if not is_patient_available(
        patient_id=target_session.patient_id,
        target_date=target_session.session_date,
        target_time=target_time,
        sessions=sessions,
        absences=absences,
        replacements=replacements,
        ignore_session_id=target_session.session_id,
    ):
        return []

    candidates: list[ReplacementCandidate] = []

    for therapist in therapists:
        if therapist.therapist_id == target_session.therapist_id:
            continue
        if target_session.robotic and not therapist.robotic_capable:
            continue
        if not is_therapist_available(
            therapist_id=therapist.therapist_id,
            target_date=target_session.session_date,
            target_time=target_time,
            sessions=sessions,
            absences=absences,
            replacements=replacements,
        ):
            continue

        workload = calculate_therapist_workload(
            therapist_id=therapist.therapist_id,
            target_date=target_session.session_date,
            sessions=sessions,
            absences=absences,
            patients=patients,
            replacements=replacements,
        )
        candidates.append(
            ReplacementCandidate(
                therapist_id=therapist.therapist_id,
                display_name=therapist.display_name,
                active_sessions=workload.active_sessions,
                infectious_sessions=workload.infectious_sessions,
                robotic_sessions=workload.robotic_sessions,
                replacement_sessions=workload.replacement_sessions,
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
    replacement_time: time | None = None,
    reason: str | None = None,
) -> ReplacementAssignment:
    """Validate and create a daily replacement overlay.

    The replacement may use the original time or a new time. Both therapist
    and patient availability are validated against the operational day. The
    base Session is never edited.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)
    target_time = replacement_time or target_session.start_time

    if any(
        replacement.target_session_id == target_session.session_id
        for replacement in replacements
    ):
        raise ValueError("Target session already has a replacement")

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

    if not is_patient_available(
        patient_id=target_session.patient_id,
        target_date=target_session.session_date,
        target_time=target_time,
        sessions=sessions,
        absences=absences,
        replacements=replacements,
        ignore_session_id=target_session.session_id,
    ):
        raise ValueError("Patient is not available")

    if not is_therapist_available(
        therapist_id=replacement_therapist_id,
        target_date=target_session.session_date,
        target_time=target_time,
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
        replacement_time=target_time,
        reason=reason,
    )
