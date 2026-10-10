from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Iterable

from .availability import is_patient_available, is_therapist_available
from .models import (
    DailyAbsence,
    DailySessionCancellation,
    Patient,
    ReplacementAssignment,
    Session,
    Therapist,
)
from .workload import calculate_therapist_workload


@dataclass(frozen=True)
class TherapistScheduleCandidate:
    therapist_id: str
    display_name: str
    active_sessions: int
    active_timeslots: int
    max_daily_timeslots: int
    available_slots: tuple[time, ...]

    @property
    def workload_label(self) -> str:
        return f"{self.display_name} {self.active_sessions}"


def _treatment_requires_robotic_capability(treatment: str) -> bool:
    text = treatment.strip().casefold()
    return text in {
        "robotic",
        "robotics",
        "robot",
        "ρομποτικό",
        "ρομποτικο",
        "ρομποτικά",
        "ρομποτικα",
    }


def _therapist_supports_treatment(therapist: Therapist, treatment: str) -> bool:
    if _treatment_requires_robotic_capability(treatment):
        return therapist.robotic_capable
    return True


def rank_therapist_candidates(
    *,
    patient_id: str,
    treatment: str,
    target_date: date,
    timeslots: Iterable[time],
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    patients: Iterable[Patient] = (),
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    ignore_session_id: str | None = None,
) -> tuple[TherapistScheduleCandidate, ...]:
    """Return feasible therapists ranked by operational load.

    The function deliberately composes the existing authoritative workload and
    availability primitives instead of reimplementing conflict rules. A
    candidate is returned only when at least one slot is simultaneously free
    for therapist and patient and the therapist still has daily timeslot
    capacity. Robotic treatments additionally require ``robotic_capable``.
    """

    therapists = tuple(therapists)
    sessions = tuple(sessions)
    patients = tuple(patients)
    absences = tuple(absences)
    replacements = tuple(replacements)
    cancellations = tuple(cancellations)
    ordered_slots = tuple(dict.fromkeys(timeslots))

    candidates: list[TherapistScheduleCandidate] = []

    for therapist in therapists:
        if not _therapist_supports_treatment(therapist, treatment):
            continue

        workload = calculate_therapist_workload(
            therapist.therapist_id,
            target_date,
            sessions=sessions,
            absences=absences,
            patients=patients,
            replacements=replacements,
            cancellations=cancellations,
        )

        if workload.active_timeslots >= therapist.max_daily_timeslots:
            continue

        available_slots = tuple(
            slot
            for slot in ordered_slots
            if is_therapist_available(
                therapist.therapist_id,
                target_date,
                slot,
                sessions=sessions,
                absences=absences,
                replacements=replacements,
                cancellations=cancellations,
            )
            and is_patient_available(
                patient_id,
                target_date,
                slot,
                sessions=sessions,
                absences=absences,
                replacements=replacements,
                cancellations=cancellations,
                ignore_session_id=ignore_session_id,
            )
        )

        if not available_slots:
            continue

        candidates.append(
            TherapistScheduleCandidate(
                therapist_id=therapist.therapist_id,
                display_name=therapist.display_name,
                active_sessions=workload.active_sessions,
                active_timeslots=workload.active_timeslots,
                max_daily_timeslots=therapist.max_daily_timeslots,
                available_slots=available_slots,
            )
        )

    candidates.sort(
        key=lambda candidate: (
            candidate.active_sessions,
            candidate.active_timeslots,
            candidate.display_name.casefold(),
            candidate.therapist_id,
        )
    )
    return tuple(candidates)
