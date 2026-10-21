from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping
from datetime import time

from .global_replacement_planner import (
    GlobalReplacementCandidate,
    GlobalReplacementPlan,
    GlobalReplacementRequest,
    plan_global_replacements,
)
from .models import (
    DailyAbsence,
    DailySessionCancellation,
    Patient,
    ReplacementAssignment,
    Session,
    Student,
    StudentAssignment,
    Therapist,
)
from .replacement_options import ReplacementProviderOption, find_replacement_options


AdjacencyPenalty = Callable[[Session, str, time], int]


def _zero_adjacency_penalty(
    target_session: Session, provider_id: str, replacement_time: time
) -> int:
    return 0


def expand_replacement_options_for_global_planner(
    target_session: Session,
    options: Iterable[ReplacementProviderOption],
    *,
    adjacency_penalty: AdjacencyPenalty | None = None,
) -> tuple[GlobalReplacementCandidate, ...]:
    """Expand provider options into provider/time candidates for global solving.

    `find_replacement_options` already performs the local hard-feasibility checks.
    This adapter preserves every feasible provider/time pair so the global solver
    can protect scarce choices instead of committing to each option's locally
    recommended time.
    """

    penalty_fn = adjacency_penalty or _zero_adjacency_penalty
    candidates: list[GlobalReplacementCandidate] = []
    for option in options:
        for slot in option.available_timeslots:
            penalty = penalty_fn(target_session, option.provider_id, slot)
            if penalty < 0:
                raise ValueError("adjacency penalty must be >= 0")
            candidates.append(
                GlobalReplacementCandidate(
                    request_id=target_session.session_id,
                    provider_id=option.provider_id,
                    provider_kind=option.provider_kind,
                    replacement_time=slot,
                    provider_active_timeslots=option.active_timeslots,
                    is_original_time=slot == target_session.start_time,
                    adjacency_penalty=penalty,
                )
            )
    return tuple(candidates)


def build_global_replacement_request(
    target_session: Session,
    options: Iterable[ReplacementProviderOption],
    *,
    adjacency_penalty: AdjacencyPenalty | None = None,
) -> GlobalReplacementRequest:
    return GlobalReplacementRequest(
        request_id=target_session.session_id,
        patient_id=target_session.patient_id,
        candidates=expand_replacement_options_for_global_planner(
            target_session,
            options,
            adjacency_penalty=adjacency_penalty,
        ),
    )


def _merge_capacity_remaining(
    capacity_by_provider: dict[str, int],
    options: Iterable[ReplacementProviderOption],
) -> None:
    for option in options:
        if option.capacity_remaining is None:
            continue
        previous = capacity_by_provider.get(option.provider_id)
        if previous is None:
            capacity_by_provider[option.provider_id] = option.capacity_remaining
        elif previous != option.capacity_remaining:
            raise ValueError(
                "capacity_remaining must be consistent for each replacement provider"
            )


def plan_replacement_sessions_globally(
    target_sessions: Iterable[Session],
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    *,
    timeslots: Iterable[time] = (),
    students: Iterable[Student] = (),
    student_assignments: Iterable[StudentAssignment] = (),
    adjacency_penalty: AdjacencyPenalty | None = None,
) -> GlobalReplacementPlan:
    """Generate real replacement options for all displaced sessions, then solve once.

    This is the first end-to-end core path for the agreed multi-absence rule:
    candidate generation remains authoritative in the existing replacement engine,
    but no local candidate is committed before every displaced session is known.
    """

    target_tuple = tuple(target_sessions)
    sessions_tuple = tuple(sessions)
    therapists_tuple = tuple(therapists)
    absences_tuple = tuple(absences)
    patients_tuple = tuple(patients)
    replacements_tuple = tuple(replacements)
    cancellations_tuple = tuple(cancellations)
    timeslots_tuple = tuple(timeslots)
    students_tuple = tuple(students)
    student_assignments_tuple = tuple(student_assignments)

    if not target_tuple:
        return GlobalReplacementPlan(assignments=(), unresolved_request_ids=())

    service_dates = {session.session_date for session in target_tuple}
    if len(service_dates) != 1:
        raise ValueError("global replacement planning requires one service date")
    if len({session.session_id for session in target_tuple}) != len(target_tuple):
        raise ValueError("target session ids must be unique")

    requests: list[GlobalReplacementRequest] = []
    capacity_by_provider: dict[str, int] = {}

    for target_session in target_tuple:
        options = find_replacement_options(
            target_session=target_session,
            therapists=therapists_tuple,
            sessions=sessions_tuple,
            absences=absences_tuple,
            patients=patients_tuple,
            replacements=replacements_tuple,
            cancellations=cancellations_tuple,
            requested_time=target_session.start_time,
            timeslots=timeslots_tuple,
            students=students_tuple,
            student_assignments=student_assignments_tuple,
        )
        _merge_capacity_remaining(capacity_by_provider, options)
        requests.append(
            build_global_replacement_request(
                target_session,
                options,
                adjacency_penalty=adjacency_penalty,
            )
        )

    return plan_global_replacements(
        requests,
        provider_capacity_remaining=capacity_by_provider,
    )
