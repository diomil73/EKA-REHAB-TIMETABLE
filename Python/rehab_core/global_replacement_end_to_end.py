from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Callable, Iterable, Mapping

from .global_replacement_integration import (
    AdjacencyPenalty,
    build_global_replacement_request,
    expand_replacement_options_for_global_planner,
)
from .global_replacement_planner import plan_global_replacements
from .models import (
    BaseScheduleEntry,
    DailyAbsence,
    DailySessionCancellation,
    Patient,
    ReplacementAssignment,
    Session,
    Student,
    StudentAssignment,
    Therapist,
)
from .provider_policy import ProviderPolicyBook
from .replacement_options import ReplacementProviderOption, find_replacement_options
from .replacement_policy import find_policy_replacement_options
from .service_hours import SpecialServiceHoursBook
from .service_replacement_options import filter_replacement_options_for_service_hours


ReplacementIdFactory = Callable[[str], str]


@dataclass(frozen=True)
class GlobalReplacementAssignmentResult:
    assignments: tuple[ReplacementAssignment, ...]
    unresolved_request_ids: tuple[str, ...]


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


def build_global_replacement_assignments(
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
    policy_book: ProviderPolicyBook | None = None,
    base_entries: Iterable[BaseScheduleEntry] = (),
    service_hours: SpecialServiceHoursBook | None = None,
    service_id_by_session: Mapping[str, str] | None = None,
    replacement_id_factory: ReplacementIdFactory,
    reason: str = "global replacement plan",
) -> GlobalReplacementAssignmentResult:
    """Build one validated global plan and materialize authoritative assignments.

    The function deliberately performs no workbook mutation. It returns normal
    ``ReplacementAssignment`` objects that can pass through the existing daily
    replacement lifecycle/ledger after the caller has reviewed and committed the
    plan. All target sessions must belong to one service date.

    When ``policy_book`` is supplied, candidate generation uses the same provider
    policy and patient cross-specialty guards as the authoritative daily workflow.
    This prevents the global optimizer from widening eligibility beyond the normal
    replacement contract merely because it is solving multiple absences together.
    """

    targets = tuple(target_sessions)
    if not targets:
        return GlobalReplacementAssignmentResult((), ())
    if len({session.session_id for session in targets}) != len(targets):
        raise ValueError("target session ids must be unique")
    if len({session.session_date for session in targets}) != 1:
        raise ValueError("global replacement planning requires one service date")

    therapists_tuple = tuple(therapists)
    sessions_tuple = tuple(sessions)
    absences_tuple = tuple(absences)
    patients_tuple = tuple(patients)
    replacements_tuple = tuple(replacements)
    cancellations_tuple = tuple(cancellations)
    timeslots_tuple = tuple(timeslots)
    students_tuple = tuple(students)
    student_assignments_tuple = tuple(student_assignments)
    base_entries_tuple = tuple(base_entries)
    service_ids = dict(service_id_by_session or {})

    requests = []
    capacity_by_provider: dict[str, int] = {}

    for target in targets:
        if policy_book is None:
            options = find_replacement_options(
                target_session=target,
                therapists=therapists_tuple,
                sessions=sessions_tuple,
                absences=absences_tuple,
                patients=patients_tuple,
                replacements=replacements_tuple,
                cancellations=cancellations_tuple,
                requested_time=target.start_time,
                timeslots=timeslots_tuple,
                students=students_tuple,
                student_assignments=student_assignments_tuple,
            )
        else:
            options = find_policy_replacement_options(
                target_session=target,
                therapists=therapists_tuple,
                sessions=sessions_tuple,
                policy_book=policy_book,
                absences=absences_tuple,
                patients=patients_tuple,
                replacements=replacements_tuple,
                cancellations=cancellations_tuple,
                requested_time=target.start_time,
                timeslots=timeslots_tuple,
                students=students_tuple,
                student_assignments=student_assignments_tuple,
                base_entries=base_entries_tuple,
            )

        if service_hours is not None:
            options = filter_replacement_options_for_service_hours(
                options,
                target_session=target,
                hours_book=service_hours,
                explicit_service_id=service_ids.get(target.session_id),
            )

        _merge_capacity_remaining(capacity_by_provider, options)
        requests.append(
            build_global_replacement_request(
                target,
                options,
                adjacency_penalty=adjacency_penalty,
            )
        )

    plan = plan_global_replacements(
        requests,
        provider_capacity_remaining=capacity_by_provider,
    )

    target_by_id = {session.session_id: session for session in targets}
    assignments = []
    seen_replacement_ids: set[str] = set()

    for chosen in plan.assignments:
        target = target_by_id[chosen.request_id]
        replacement_id = replacement_id_factory(target.session_id)
        if not replacement_id:
            raise ValueError("replacement_id_factory returned an empty id")
        if replacement_id in seen_replacement_ids:
            raise ValueError("replacement_id_factory returned a duplicate id")
        seen_replacement_ids.add(replacement_id)

        assignments.append(
            ReplacementAssignment(
                replacement_id=replacement_id,
                target_session_id=target.session_id,
                patient_id=target.patient_id,
                original_therapist_id=target.therapist_id,
                replacement_therapist_id=chosen.provider_id,
                replacement_date=target.session_date,
                replacement_time=chosen.replacement_time,
                reason=reason,
                replacement_provider_kind=chosen.provider_kind,
            )
        )

    return GlobalReplacementAssignmentResult(
        assignments=tuple(sorted(assignments, key=lambda item: item.target_session_id)),
        unresolved_request_ids=plan.unresolved_request_ids,
    )
