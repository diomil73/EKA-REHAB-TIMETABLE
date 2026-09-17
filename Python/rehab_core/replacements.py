from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable

from .availability import is_therapist_available
from .models import AbsenceKind, DailyAbsence, Patient, Session, Therapist


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
    patient_by_id: dict[str, Patient],
) -> tuple[int, int]:
    """Return (active sessions, infectious active sessions) for the day.

    Patient absences remove the corresponding base session from today's
    operational workload. The base schedule itself remains unchanged.
    """

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

    return active_sessions, infectious_sessions


def find_replacement_candidates(
    target_session: Session,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
) -> list[ReplacementCandidate]:
    """Return available replacement therapists in a deterministic order.

    Milestone 2 rules:
    - never suggest the therapist already assigned to the base session;
    - exclude therapists absent at the requested date/time;
    - exclude therapists busy at that date/time, except when their own patient
      is absent and therefore their slot is operationally free;
    - robotic sessions require a robotic-capable therapist;
    - balance workload using today's active sessions;
    - for an infectious target patient, prefer the lower infectious workload
      before total workload;
    - do not mutate the base schedule.

    Day-pattern compatibility and persisted daily replacement overlays will be
    added as separate milestones rather than hidden inside this first engine.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    patients = tuple(patients)
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
        ):
            continue

        active_sessions, infectious_sessions = _daily_workload(
            therapist_id=therapist.therapist_id,
            target_date=target_session.session_date,
            sessions=sessions,
            absences=absences,
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
