from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from .daily_state import DailySessionState, DailySessionStatus
from .models import SessionCancellationKind


@dataclass(frozen=True)
class DailyTreatmentStatistics:
    scheduled_sessions: int
    performed_sessions: int
    replaced_sessions: int
    patient_absent_sessions: int
    cancelled_sessions: int
    department_postponed_sessions: int
    patient_no_show_sessions: int
    treatment_suspended_sessions: int
    unresolved_sessions: int

    @property
    def non_performed_sessions(self) -> int:
        return self.patient_absent_sessions + self.cancelled_sessions

    @property
    def accounted_sessions(self) -> int:
        return (
            self.performed_sessions
            + self.non_performed_sessions
            + self.unresolved_sessions
        )

    @property
    def is_finalized(self) -> bool:
        """Whether every scheduled session has a final operational outcome."""

        return self.unresolved_sessions == 0


def calculate_daily_treatment_statistics(
    states: Iterable[DailySessionState],
) -> DailyTreatmentStatistics:
    """Summarize concrete-date treatment outcomes from effective daily state.

    ACTIVE and REPLACED sessions count as performed. PATIENT_ABSENT and
    CANCELLED sessions count as explicit non-performed outcomes. A remaining
    THERAPIST_ABSENT state is deliberately reported as unresolved rather than
    silently counted as a final cancellation; it must still be placed or
    explicitly suspended before the day is finalized.
    """

    state_list = tuple(states)
    performed = sum(
        state.status in {DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED}
        for state in state_list
    )
    replaced = sum(
        state.status == DailySessionStatus.REPLACED for state in state_list
    )
    patient_absent = sum(
        state.status == DailySessionStatus.PATIENT_ABSENT for state in state_list
    )
    cancelled = sum(
        state.status == DailySessionStatus.CANCELLED for state in state_list
    )
    unresolved = sum(
        state.status == DailySessionStatus.THERAPIST_ABSENT for state in state_list
    )

    def cancelled_kind(kind: SessionCancellationKind) -> int:
        return sum(
            state.status == DailySessionStatus.CANCELLED
            and state.cancellation_kind == kind
            for state in state_list
        )

    return DailyTreatmentStatistics(
        scheduled_sessions=len(state_list),
        performed_sessions=performed,
        replaced_sessions=replaced,
        patient_absent_sessions=patient_absent,
        cancelled_sessions=cancelled,
        department_postponed_sessions=cancelled_kind(
            SessionCancellationKind.DEPARTMENT_POSTPONED
        ),
        patient_no_show_sessions=cancelled_kind(
            SessionCancellationKind.PATIENT_NO_SHOW
        ),
        treatment_suspended_sessions=cancelled_kind(
            SessionCancellationKind.TREATMENT_SUSPENDED
        ),
        unresolved_sessions=unresolved,
    )
