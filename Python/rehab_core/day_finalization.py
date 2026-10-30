from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from enum import Enum
from typing import Iterable

from .daily_state import DailySessionState
from .daily_statistics import calculate_daily_treatment_statistics


class FinalizedSnapshotStatus(str, Enum):
    CURRENT_FINAL = "CURRENT_FINAL"
    SUPERSEDED = "SUPERSEDED"


class DayFinalizationError(ValueError):
    """Raised when the operational day is not safe to finalize."""


@dataclass(frozen=True)
class FinalizedDaySnapshot:
    """Immutable historical header for one successfully finalized service day.

    This is deliberately a pure core model. Excel persistence/version-chain
    mutation is handled by a later integration layer; this object captures the
    validated facts that are allowed to enter that history layer.
    """

    snapshot_id: str
    service_date: date
    version: int
    status: FinalizedSnapshotStatus
    finalized_at: datetime
    supersedes_snapshot_id: str | None
    unresolved_sessions: int
    planned_session_count: int | None
    final_session_count: int
    patient_absence_count: int
    cancellation_count: int
    no_show_count: int
    replacement_count: int
    treatment_suspension_count: int
    approved_exception_count: int


def build_finalized_day_snapshot(
    states: Iterable[DailySessionState],
    *,
    snapshot_id: str,
    service_date: date,
    finalized_at: datetime,
    version: int = 1,
    supersedes_snapshot_id: str | None = None,
    planned_session_count: int | None = None,
    approved_exception_count: int = 0,
) -> FinalizedDaySnapshot:
    """Validate a concrete operational day and freeze its final outcome counts.

    Finalization is fail-closed. Every supplied state must belong to the target
    service date and no session may remain unresolved. The snapshot is built
    only from the effective daily truth, never from the mutable recurring plan.
    """

    snapshot_id = snapshot_id.strip()
    if not snapshot_id:
        raise DayFinalizationError("snapshot_id is required")
    if version < 1:
        raise DayFinalizationError("snapshot version must be at least 1")
    if finalized_at.tzinfo is None or finalized_at.utcoffset() is None:
        raise DayFinalizationError("finalized_at must include a timezone offset")
    if planned_session_count is not None and planned_session_count < 0:
        raise DayFinalizationError("planned_session_count cannot be negative")
    if approved_exception_count < 0:
        raise DayFinalizationError("approved_exception_count cannot be negative")

    state_list = tuple(states)
    wrong_date_ids = tuple(
        state.session_id for state in state_list if state.session_date != service_date
    )
    if wrong_date_ids:
        joined = ", ".join(wrong_date_ids)
        raise DayFinalizationError(
            f"Cannot finalize {service_date.isoformat()}: states from another date: {joined}"
        )

    stats = calculate_daily_treatment_statistics(state_list)
    if not stats.is_finalized:
        unresolved_ids = tuple(
            state.session_id for state in state_list if state.needs_replacement
        )
        joined = ", ".join(unresolved_ids) or str(stats.unresolved_sessions)
        raise DayFinalizationError(
            f"Cannot finalize {service_date.isoformat()}: unresolved sessions: {joined}"
        )

    return FinalizedDaySnapshot(
        snapshot_id=snapshot_id,
        service_date=service_date,
        version=version,
        status=FinalizedSnapshotStatus.CURRENT_FINAL,
        finalized_at=finalized_at,
        supersedes_snapshot_id=supersedes_snapshot_id,
        unresolved_sessions=stats.unresolved_sessions,
        planned_session_count=planned_session_count,
        final_session_count=stats.performed_sessions,
        patient_absence_count=stats.patient_absent_sessions,
        cancellation_count=stats.cancelled_sessions,
        no_show_count=stats.patient_no_show_sessions,
        replacement_count=stats.replaced_sessions,
        treatment_suspension_count=stats.treatment_suspended_sessions,
        approved_exception_count=approved_exception_count,
    )
