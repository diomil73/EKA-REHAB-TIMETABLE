from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from rehab_core.global_replacement_end_to_end import GlobalReplacementAssignmentResult
from rehab_core.models import ReplacementAssignment, Session

from .daily_replacement_facts import (
    DailyReplacementFactRow,
    append_daily_replacement_facts,
    build_daily_replacement_fact_rows,
)


class GlobalReplacementLifecycleError(RuntimeError):
    """Raised when a global plan conflicts with the active replacement lifecycle."""


@dataclass(frozen=True)
class GlobalReplacementFactBatch:
    fact_rows: tuple[DailyReplacementFactRow, ...]
    unresolved_request_ids: tuple[str, ...]
    assignments_to_persist: tuple[ReplacementAssignment, ...] = ()
    already_active_replacement_ids: tuple[str, ...] = ()


@dataclass(frozen=True)
class GlobalReplacementPersistenceReport:
    requested_assignment_count: int
    persisted_fact_count: int
    persisted_replacement_ids: tuple[str, ...]
    already_active_replacement_ids: tuple[str, ...]
    unresolved_request_ids: tuple[str, ...]


PersistenceService = Callable[..., int]


def _semantic_assignment_key(value: ReplacementAssignment) -> tuple[object, ...]:
    return (
        value.target_session_id,
        value.patient_id,
        value.original_therapist_id,
        value.replacement_therapist_id,
        value.replacement_date,
        value.replacement_time.replace(second=0, microsecond=0),
        value.replacement_provider_kind,
    )


def prepare_global_replacement_fact_batch(
    result: GlobalReplacementAssignmentResult,
    *,
    sessions: Iterable[Session],
    existing_active_replacements: Iterable[ReplacementAssignment] = (),
) -> GlobalReplacementFactBatch:
    """Validate a global plan against the existing authoritative fact lifecycle.

    The bridge is mutation-free. New resolved assignments are converted through the
    existing ``DAILY_REPLACEMENT_FACTS`` schema validator, while unresolved request
    ids remain explicit for the orphan-resolution/finalization workflow.

    If the same target session/date already has an ACTIVE assignment, the operation
    is allowed only when the assignment is semantically identical and uses the same
    ReplacementId. That case is treated as idempotent and no duplicate fact row is
    produced. Any different ACTIVE assignment is rejected before workbook mutation.
    """

    session_tuple = tuple(sessions)
    existing_tuple = tuple(existing_active_replacements)

    existing_by_target: dict[tuple[str, object], ReplacementAssignment] = {}
    existing_ids: dict[str, ReplacementAssignment] = {}
    for existing in existing_tuple:
        target_key = (existing.target_session_id, existing.replacement_date)
        if target_key in existing_by_target:
            raise GlobalReplacementLifecycleError(
                f"Target session/date {target_key!r} already has multiple active replacements"
            )
        if existing.replacement_id in existing_ids:
            raise GlobalReplacementLifecycleError(
                f"Duplicate active ReplacementId {existing.replacement_id!r}"
            )
        existing_by_target[target_key] = existing
        existing_ids[existing.replacement_id] = existing

    new_assignments: list[ReplacementAssignment] = []
    already_active_ids: list[str] = []

    for assignment in result.assignments:
        target_key = (assignment.target_session_id, assignment.replacement_date)
        existing = existing_by_target.get(target_key)
        if existing is not None:
            if (
                existing.replacement_id == assignment.replacement_id
                and _semantic_assignment_key(existing) == _semantic_assignment_key(assignment)
            ):
                already_active_ids.append(existing.replacement_id)
                continue
            raise GlobalReplacementLifecycleError(
                f"Target session {assignment.target_session_id!r} already has an active replacement"
            )

        old_id = existing_ids.get(assignment.replacement_id)
        if old_id is not None:
            raise GlobalReplacementLifecycleError(
                f"ReplacementId {assignment.replacement_id!r} is already active for another target"
            )

        new_assignments.append(assignment)

    try:
        fact_rows = build_daily_replacement_fact_rows(
            new_assignments,
            sessions=session_tuple,
        )
    except Exception as exc:
        raise GlobalReplacementLifecycleError(str(exc)) from exc

    return GlobalReplacementFactBatch(
        fact_rows=fact_rows,
        unresolved_request_ids=tuple(result.unresolved_request_ids),
        assignments_to_persist=tuple(new_assignments),
        already_active_replacement_ids=tuple(sorted(already_active_ids)),
    )


def persist_global_replacement_result(
    workbook_path: str | Path,
    result: GlobalReplacementAssignmentResult,
    *,
    sessions: Iterable[Session],
    existing_active_replacements: Iterable[ReplacementAssignment] = (),
    require_fully_resolved: bool = True,
    persistence_service: PersistenceService = append_daily_replacement_facts,
) -> GlobalReplacementPersistenceReport:
    """Persist one validated global plan through the existing authoritative fact store.

    This is the controlled transaction boundary between the global planner and
    ``DAILY_REPLACEMENT_FACTS``. By default, any unresolved request blocks the
    write completely. Callers may opt out only for non-final preview workflows.

    Existing semantically identical ACTIVE facts are idempotent and are not sent
    back to the writer. Conflicts are rejected by the mutation-free preparation
    step before ``persistence_service`` is called.
    """

    session_tuple = tuple(sessions)
    batch = prepare_global_replacement_fact_batch(
        result,
        sessions=session_tuple,
        existing_active_replacements=existing_active_replacements,
    )

    if require_fully_resolved and batch.unresolved_request_ids:
        unresolved = ", ".join(batch.unresolved_request_ids)
        raise GlobalReplacementLifecycleError(
            "Global replacement plan is not finalized; unresolved request(s): " + unresolved
        )

    if batch.assignments_to_persist:
        try:
            written = persistence_service(
                workbook_path,
                batch.assignments_to_persist,
                sessions=session_tuple,
            )
        except GlobalReplacementLifecycleError:
            raise
        except Exception as exc:
            raise GlobalReplacementLifecycleError(str(exc)) from exc
    else:
        written = 0

    return GlobalReplacementPersistenceReport(
        requested_assignment_count=len(result.assignments),
        persisted_fact_count=written,
        persisted_replacement_ids=tuple(
            assignment.replacement_id for assignment in batch.assignments_to_persist
        ),
        already_active_replacement_ids=batch.already_active_replacement_ids,
        unresolved_request_ids=batch.unresolved_request_ids,
    )
