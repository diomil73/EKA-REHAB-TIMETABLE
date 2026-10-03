from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from .models import (
    DailyAbsence,
    DailySessionCancellation,
    Patient,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
    Student,
    StudentAssignment,
    Therapist,
)
from .replacements import ReplacementCandidate, find_replacement_candidates


def _minutes(value: time) -> int:
    return value.hour * 60 + value.minute


def order_timeslots_by_preference(
    timeslots: Iterable[time], requested_time: time
) -> tuple[time, ...]:
    unique = set(timeslots)
    return tuple(
        sorted(
            unique,
            key=lambda slot: (
                0 if slot == requested_time else 1,
                abs(_minutes(slot) - _minutes(requested_time)),
                slot,
            ),
        )
    )


@dataclass(frozen=True)
class ReplacementProviderOption:
    provider_id: str
    display_name: str
    provider_kind: ReplacementProviderKind
    active_sessions: int
    infectious_sessions: int
    robotic_sessions: int
    replacement_sessions: int
    requested_time: time
    recommended_time: time
    exact_time_available: bool
    available_timeslots: tuple[time, ...]
    capacity_limit: int | None = None
    capacity_remaining: int | None = None

    @property
    def uses_alternative_time(self) -> bool:
        return self.recommended_time != self.requested_time


def _to_option(
    candidate: ReplacementCandidate, requested_time: time
) -> ReplacementProviderOption:
    ordered = order_timeslots_by_preference(
        candidate.available_timeslots, requested_time
    )
    if not ordered:
        raise ValueError("Replacement candidate has no available timeslots")
    return ReplacementProviderOption(
        provider_id=candidate.provider_id,
        display_name=candidate.display_name,
        provider_kind=candidate.provider_kind,
        active_sessions=candidate.active_sessions,
        infectious_sessions=candidate.infectious_sessions,
        robotic_sessions=candidate.robotic_sessions,
        replacement_sessions=candidate.replacement_sessions,
        requested_time=requested_time,
        recommended_time=ordered[0],
        exact_time_available=candidate.exact_time_available,
        available_timeslots=ordered,
        capacity_limit=candidate.capacity_limit,
        capacity_remaining=candidate.capacity_remaining,
    )


def find_replacement_options(
    target_session: Session,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    *,
    requested_time: time | None = None,
    timeslots: Iterable[time] = (),
    students: Iterable[Student] = (),
    student_assignments: Iterable[StudentAssignment] = (),
) -> list[ReplacementProviderOption]:
    """Return ranked providers together with a recommended feasible timeslot."""

    preferred_time = requested_time or target_session.start_time
    candidates = find_replacement_candidates(
        target_session=target_session,
        therapists=therapists,
        sessions=sessions,
        absences=absences,
        patients=patients,
        replacements=replacements,
        cancellations=cancellations,
        replacement_time=preferred_time,
        timeslots=timeslots,
        students=students,
        student_assignments=student_assignments,
    )
    return [_to_option(candidate, preferred_time) for candidate in candidates]
