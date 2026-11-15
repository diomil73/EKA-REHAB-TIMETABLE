from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable, Mapping

from .daily_state import DailySessionState, DailySessionStatus
from .models import ReplacementProviderKind, SessionCancellationKind


class TherapistDayHistoryError(RuntimeError):
    """Raised when finalized therapist/day workload cannot be derived safely."""


@dataclass(frozen=True)
class TherapistDayHistoryRow:
    snapshot_id: str
    service_date: date
    therapist_id: str
    capacity_timeslots: int | None
    planned_occupied_timeslots: int
    planned_sessions: int
    final_occupied_timeslots: int
    final_active_sessions: int
    replacements_received: int
    released_by_cancellation_or_no_show: int
    treatment_suspensions_affecting_load: int
    approved_shared_or_over_capacity_events: int = 0


def build_therapist_day_history_rows(
    states: Iterable[DailySessionState],
    *,
    snapshot_id: str,
    service_date: date,
    therapist_ids: Iterable[str] = (),
    capacity_timeslots_by_therapist: Mapping[str, int] | None = None,
    approved_exception_events_by_therapist: Mapping[str, int] | None = None,
) -> tuple[TherapistDayHistoryRow, ...]:
    """Freeze therapist/day workload from one finalized effective daily state.

    Planned metrics use original therapist/time ownership. Final workload uses
    effective therapist/time after replacements and cancellations. Occupied
    timeslots count unique clock slots, while active sessions count patient
    sessions. A legal shared slot can therefore be one occupied timeslot and
    multiple active sessions.

    Student replacement work is deliberately excluded from therapist rows and
    will be persisted through the student-history path.
    """

    if not snapshot_id.strip():
        raise TherapistDayHistoryError("snapshot_id cannot be blank")

    state_list = tuple(states)
    wrong_date = tuple(
        state.session_id for state in state_list if state.session_date != service_date
    )
    if wrong_date:
        raise TherapistDayHistoryError(
            "Therapist/day history received states from another service date: "
            + ", ".join(wrong_date)
        )

    unresolved = tuple(
        state.session_id
        for state in state_list
        if state.status == DailySessionStatus.THERAPIST_ABSENT
    )
    if unresolved:
        raise TherapistDayHistoryError(
            "Cannot freeze therapist/day history with unresolved sessions: "
            + ", ".join(unresolved)
        )

    capacity = dict(capacity_timeslots_by_therapist or {})
    approved = dict(approved_exception_events_by_therapist or {})
    if any(value < 0 for value in capacity.values()):
        raise TherapistDayHistoryError("capacity_timeslots cannot be negative")
    if any(value < 0 for value in approved.values()):
        raise TherapistDayHistoryError("approved exception counts cannot be negative")

    provider_ids = {str(value).strip() for value in therapist_ids if str(value).strip()}
    provider_ids.update(
        state.original_therapist_id.strip()
        for state in state_list
        if state.original_therapist_id.strip()
    )
    provider_ids.update(
        state.effective_therapist_id.strip()
        for state in state_list
        if state.effective_therapist_id
        and state.effective_provider_kind != ReplacementProviderKind.STUDENT
    )
    provider_ids.update(capacity)
    provider_ids.update(approved)

    rows: list[TherapistDayHistoryRow] = []
    for therapist_id in sorted(provider_ids, key=str.casefold):
        planned = tuple(
            state for state in state_list if state.original_therapist_id == therapist_id
        )
        performed = tuple(
            state
            for state in state_list
            if state.status in {DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED}
            and state.effective_therapist_id == therapist_id
            and state.effective_provider_kind != ReplacementProviderKind.STUDENT
        )
        planned_slots = {(state.original_time) for state in planned}
        final_slots = {
            state.effective_time
            for state in performed
            if state.effective_time is not None
        }
        released = sum(
            state.status == DailySessionStatus.PATIENT_ABSENT
            or (
                state.status == DailySessionStatus.CANCELLED
                and state.cancellation_kind
                in {
                    SessionCancellationKind.PATIENT_NO_SHOW,
                    SessionCancellationKind.DEPARTMENT_POSTPONED,
                }
            )
            for state in planned
        )
        suspensions = sum(
            state.status == DailySessionStatus.CANCELLED
            and state.cancellation_kind == SessionCancellationKind.TREATMENT_SUSPENDED
            for state in planned
        )
        replacements_received = sum(
            state.status == DailySessionStatus.REPLACED for state in performed
        )

        rows.append(
            TherapistDayHistoryRow(
                snapshot_id=snapshot_id,
                service_date=service_date,
                therapist_id=therapist_id,
                capacity_timeslots=capacity.get(therapist_id),
                planned_occupied_timeslots=len(planned_slots),
                planned_sessions=len(planned),
                final_occupied_timeslots=len(final_slots),
                final_active_sessions=len(performed),
                replacements_received=replacements_received,
                released_by_cancellation_or_no_show=released,
                treatment_suspensions_affecting_load=suspensions,
                approved_shared_or_over_capacity_events=approved.get(therapist_id, 0),
            )
        )

    return tuple(rows)
