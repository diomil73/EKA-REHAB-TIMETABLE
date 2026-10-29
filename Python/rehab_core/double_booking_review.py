from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from .double_booking_exception import (
    DOUBLE_BOOKING_WARNING,
    DoubleBookingExceptionProposal,
    propose_double_booking_exception,
)
from .global_replacement_end_to_end import GlobalReplacementAssignmentResult
from .models import (
    BaseScheduleEntry,
    DailyAbsence,
    DailySessionCancellation,
    Patient,
    ReplacementAssignment,
    Session,
    Therapist,
)
from .provider_policy import ProviderPolicyBook


@dataclass(frozen=True)
class DoubleBookingReviewOption:
    therapist_id: str
    therapist_name: str
    replacement_time: time
    warning: str


@dataclass(frozen=True)
class DoubleBookingReviewItem:
    target_session_id: str
    patient_id: str
    patient_name: str
    original_therapist_id: str
    original_time: time
    options: tuple[DoubleBookingReviewOption, ...]

    @property
    def has_options(self) -> bool:
        return bool(self.options)


def build_double_booking_review_items(
    result: GlobalReplacementAssignmentResult,
    *,
    sessions: Iterable[Session],
    therapists: Iterable[Therapist],
    patients: Iterable[Patient] = (),
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    timeslots: Iterable[time] = (),
    policy_book: ProviderPolicyBook | None = None,
    base_entries: Iterable[BaseScheduleEntry] = (),
) -> tuple[DoubleBookingReviewItem, ...]:
    """Build UI-ready double-booking choices only for unresolved normal-plan sessions.

    This function never creates or confirms an assignment. It probes the existing
    explicit exception validator and exposes only choices that may legally be
    offered to the user for manual confirmation.
    """

    unresolved_ids = tuple(result.unresolved_request_ids)
    if not unresolved_ids:
        return ()

    sessions_tuple = tuple(sessions)
    therapists_tuple = tuple(therapists)
    patients_tuple = tuple(patients)
    absences_tuple = tuple(absences)
    replacements_tuple = tuple(replacements) + tuple(result.assignments)
    cancellations_tuple = tuple(cancellations)
    base_entries_tuple = tuple(base_entries)
    explicit_timeslots = tuple(timeslots)

    session_by_id = {session.session_id: session for session in sessions_tuple}
    patient_name_by_id = {
        patient.patient_id: patient.display_name for patient in patients_tuple
    }

    items: list[DoubleBookingReviewItem] = []
    for session_id in unresolved_ids:
        target = session_by_id.get(session_id)
        if target is None:
            raise ValueError(f"Unknown unresolved session id: {session_id}")

        candidate_times = explicit_timeslots or (target.start_time,)
        options: list[DoubleBookingReviewOption] = []

        for therapist in therapists_tuple:
            if therapist.therapist_id == target.therapist_id:
                continue
            for candidate_time in candidate_times:
                try:
                    proposal: DoubleBookingExceptionProposal = (
                        propose_double_booking_exception(
                            target_session=target,
                            unresolved_session_ids=unresolved_ids,
                            therapist_id=therapist.therapist_id,
                            replacement_time=candidate_time,
                            therapists=therapists_tuple,
                            sessions=sessions_tuple,
                            absences=absences_tuple,
                            patients=patients_tuple,
                            replacements=replacements_tuple,
                            cancellations=cancellations_tuple,
                            policy_book=policy_book,
                            base_entries=base_entries_tuple,
                        )
                    )
                except ValueError:
                    continue

                options.append(
                    DoubleBookingReviewOption(
                        therapist_id=therapist.therapist_id,
                        therapist_name=therapist.display_name,
                        replacement_time=proposal.replacement_time,
                        warning=proposal.warning,
                    )
                )

        options.sort(
            key=lambda item: (
                0 if item.replacement_time == target.start_time else 1,
                item.replacement_time,
                item.therapist_name.casefold(),
                item.therapist_id,
            )
        )
        items.append(
            DoubleBookingReviewItem(
                target_session_id=target.session_id,
                patient_id=target.patient_id,
                patient_name=patient_name_by_id.get(target.patient_id, target.patient_id),
                original_therapist_id=target.therapist_id,
                original_time=target.start_time,
                options=tuple(options),
            )
        )

    return tuple(items)


__all__ = [
    "DOUBLE_BOOKING_WARNING",
    "DoubleBookingReviewItem",
    "DoubleBookingReviewOption",
    "build_double_booking_review_items",
]
