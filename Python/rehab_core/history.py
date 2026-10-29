from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Iterable

from .daily_state import DailySessionState, DailySessionStatus
from .daily_statistics import DailyTreatmentStatistics, calculate_daily_treatment_statistics


class DayFinalizationError(ValueError):
    """Raised when a concrete operational day is not safe to finalize."""

    def __init__(
        self,
        *,
        target_date: date,
        unresolved_session_ids: tuple[str, ...] = (),
        orphan_session_ids: tuple[str, ...] = (),
        wrong_date_session_ids: tuple[str, ...] = (),
    ) -> None:
        self.target_date = target_date
        self.unresolved_session_ids = unresolved_session_ids
        self.orphan_session_ids = orphan_session_ids
        self.wrong_date_session_ids = wrong_date_session_ids

        problems: list[str] = []
        if wrong_date_session_ids:
            problems.append(
                "states from another date: " + ", ".join(wrong_date_session_ids)
            )
        if unresolved_session_ids:
            problems.append(
                "unresolved sessions: " + ", ".join(unresolved_session_ids)
            )
        if orphan_session_ids:
            problems.append("orphan sessions: " + ", ".join(orphan_session_ids))

        detail = "; ".join(problems) or "day is not finalizable"
        super().__init__(f"Cannot finalize {target_date.isoformat()}: {detail}")


@dataclass(frozen=True)
class FinalizedSessionFact:
    """Immutable historical fact derived from one effective daily session state."""

    session_id: str
    patient_id: str
    session_date: date
    status: DailySessionStatus
    original_therapist_id: str
    original_time: time
    effective_therapist_id: str | None
    effective_time: time | None
    replacement_id: str | None
    cancellation_id: str | None
    cancellation_kind: str | None
    reason: str | None


@dataclass(frozen=True)
class TherapistDayHistory:
    """Historical therapist/day facts with planned and actual workload separated.

    Planned workload is derived from the immutable original assignment. Actual
    workload is derived from the final effective provider/time, so a replacement
    increases the replacement provider's actual workload without rewriting the
    original therapist's planned programme.
    """

    therapist_id: str
    target_date: date
    planned_session_count: int
    actual_session_count: int
    occupied_timeslots: tuple[time, ...]
    active_session_count: int
    replacement_session_count: int
    patient_absent_session_count: int
    cancelled_session_count: int


@dataclass(frozen=True)
class FinalizedDaySnapshot:
    """Versionable, workbook-independent historical representation of one day."""

    target_date: date
    version: int
    treatment_statistics: DailyTreatmentStatistics
    sessions: tuple[FinalizedSessionFact, ...]
    therapist_history: tuple[TherapistDayHistory, ...]


def validate_day_can_be_finalized(
    target_date: date,
    states: Iterable[DailySessionState],
    *,
    orphan_session_ids: Iterable[str] = (),
) -> tuple[DailySessionState, ...]:
    """Validate the close-day gate without mutating operational state.

    A day cannot close while a therapist-absence session is still unresolved,
    while an orphan session remains outside the normal daily-state model, or
    when the caller accidentally mixes states from another date.
    """

    state_list = tuple(states)
    wrong_date = tuple(
        state.session_id
        for state in state_list
        if state.session_date != target_date
    )
    unresolved = tuple(
        state.session_id
        for state in state_list
        if state.status == DailySessionStatus.THERAPIST_ABSENT
    )
    orphans = tuple(dict.fromkeys(str(value) for value in orphan_session_ids if value))

    if wrong_date or unresolved or orphans:
        raise DayFinalizationError(
            target_date=target_date,
            unresolved_session_ids=unresolved,
            orphan_session_ids=orphans,
            wrong_date_session_ids=wrong_date,
        )

    return state_list


def _session_fact(state: DailySessionState) -> FinalizedSessionFact:
    cancellation_kind = (
        state.cancellation_kind.value if state.cancellation_kind is not None else None
    )
    return FinalizedSessionFact(
        session_id=state.session_id,
        patient_id=state.patient_id,
        session_date=state.session_date,
        status=state.status,
        original_therapist_id=state.original_therapist_id,
        original_time=state.original_time,
        effective_therapist_id=state.effective_therapist_id,
        effective_time=state.effective_time,
        replacement_id=state.replacement_id,
        cancellation_id=state.cancellation_id,
        cancellation_kind=cancellation_kind,
        reason=state.reason,
    )


def build_therapist_day_history(
    target_date: date,
    states: Iterable[DailySessionState],
) -> tuple[TherapistDayHistory, ...]:
    """Build deterministic therapist/day history from final effective state."""

    state_list = tuple(states)
    therapist_ids = sorted(
        {
            therapist_id
            for state in state_list
            for therapist_id in (
                state.original_therapist_id,
                state.effective_therapist_id,
            )
            if therapist_id
        }
    )

    rows: list[TherapistDayHistory] = []
    for therapist_id in therapist_ids:
        original_states = tuple(
            state
            for state in state_list
            if state.original_therapist_id == therapist_id
        )
        actual_states = tuple(
            state
            for state in state_list
            if state.effective_therapist_id == therapist_id
            and state.effective_time is not None
            and state.status in {DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED}
        )
        occupied = tuple(sorted({state.effective_time for state in actual_states}))

        rows.append(
            TherapistDayHistory(
                therapist_id=therapist_id,
                target_date=target_date,
                planned_session_count=len(original_states),
                actual_session_count=len(actual_states),
                occupied_timeslots=occupied,
                active_session_count=sum(
                    state.status == DailySessionStatus.ACTIVE
                    for state in actual_states
                ),
                replacement_session_count=sum(
                    state.status == DailySessionStatus.REPLACED
                    for state in actual_states
                ),
                patient_absent_session_count=sum(
                    state.status == DailySessionStatus.PATIENT_ABSENT
                    for state in original_states
                ),
                cancelled_session_count=sum(
                    state.status == DailySessionStatus.CANCELLED
                    for state in original_states
                ),
            )
        )

    return tuple(rows)


def build_finalized_day_snapshot(
    target_date: date,
    states: Iterable[DailySessionState],
    *,
    version: int = 1,
    orphan_session_ids: Iterable[str] = (),
) -> FinalizedDaySnapshot:
    """Create a deterministic finalized snapshot without workbook mutation."""

    if version < 1:
        raise ValueError("Snapshot version must be >= 1")

    state_list = validate_day_can_be_finalized(
        target_date,
        states,
        orphan_session_ids=orphan_session_ids,
    )
    statistics = calculate_daily_treatment_statistics(state_list)
    if not statistics.is_finalized:
        raise DayFinalizationError(target_date=target_date)

    ordered_states = tuple(
        sorted(
            state_list,
            key=lambda state: (
                state.original_time,
                state.patient_id,
                state.session_id,
            ),
        )
    )

    return FinalizedDaySnapshot(
        target_date=target_date,
        version=version,
        treatment_statistics=statistics,
        sessions=tuple(_session_fact(state) for state in ordered_states),
        therapist_history=build_therapist_day_history(target_date, state_list),
    )
