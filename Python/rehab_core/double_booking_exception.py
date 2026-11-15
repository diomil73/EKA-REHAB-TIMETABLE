from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from .availability import is_patient_available, is_therapist_available
from .models import (
    AbsenceKind,
    BaseScheduleEntry,
    DailyAbsence,
    DailySessionCancellation,
    Patient,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
    Therapist,
)
from .patient_schedule import validate_patient_cross_specialty_time
from .provider_policy import ProviderPolicyBook
from .replacement_policy import validate_policy_replacement_time


DOUBLE_BOOKING_WARNING = (
    "Ο επιλεγμένος θεραπευτής έχει ήδη ενεργή συνεδρία στην ίδια ώρα. "
    "Η ενέργεια θα δημιουργήσει ρητή διπλοκράτηση και πρέπει να επιβεβαιωθεί."
)


@dataclass(frozen=True)
class DoubleBookingExceptionProposal:
    target_session_id: str
    patient_id: str
    therapist_id: str
    replacement_time: time
    warning: str = DOUBLE_BOOKING_WARNING


def _session_cancelled(
    session: Session,
    cancellations: tuple[DailySessionCancellation, ...],
) -> bool:
    return any(cancellation.applies_to(session) for cancellation in cancellations)


def _patient_absent(
    patient_id: str,
    *,
    target_session: Session,
    target_time: time,
    absences: tuple[DailyAbsence, ...],
) -> bool:
    return any(
        absence.absence_kind == AbsenceKind.PATIENT
        and absence.subject_id == patient_id
        and absence.covers(target_session.session_date, target_time)
        for absence in absences
    )


def _therapist_absent(
    therapist_id: str,
    *,
    target_session: Session,
    target_time: time,
    absences: tuple[DailyAbsence, ...],
) -> bool:
    return any(
        absence.absence_kind == AbsenceKind.THERAPIST
        and absence.subject_id == therapist_id
        and absence.covers(target_session.session_date, target_time)
        for absence in absences
    )


def _has_live_therapist_slot_conflict(
    therapist_id: str,
    *,
    target_session: Session,
    target_time: time,
    sessions: tuple[Session, ...],
    absences: tuple[DailyAbsence, ...],
    replacements: tuple[ReplacementAssignment, ...],
    cancellations: tuple[DailySessionCancellation, ...],
) -> bool:
    replaced_session_ids = {item.target_session_id for item in replacements}

    for replacement in replacements:
        if replacement.replacement_provider_kind != ReplacementProviderKind.THERAPIST:
            continue
        if (
            replacement.replacement_therapist_id == therapist_id
            and replacement.replacement_date == target_session.session_date
            and replacement.replacement_time == target_time
            and not _patient_absent(
                replacement.patient_id,
                target_session=target_session,
                target_time=target_time,
                absences=absences,
            )
        ):
            return True

    for session in sessions:
        if session.session_id == target_session.session_id:
            continue
        if (
            session.therapist_id != therapist_id
            or session.session_date != target_session.session_date
            or session.start_time != target_time
        ):
            continue
        if _session_cancelled(session, cancellations):
            continue
        if session.session_id in replaced_session_ids:
            continue
        if _patient_absent(
            session.patient_id,
            target_session=target_session,
            target_time=target_time,
            absences=absences,
        ):
            continue
        return True

    return False


def propose_double_booking_exception(
    *,
    target_session: Session,
    unresolved_session_ids: Iterable[str],
    therapist_id: str,
    replacement_time: time,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    policy_book: ProviderPolicyBook | None = None,
    base_entries: Iterable[BaseScheduleEntry] = (),
) -> DoubleBookingExceptionProposal:
    """Validate one user-selected double-booking exception proposal.

    This deliberately relaxes only therapist same-slot occupancy. Every other
    hard guard remains active: the target must already be unresolved by the
    normal global solver, the patient must be free, the therapist must not be
    absent, provider policy must allow the selected time, and cross-specialty
    patient conflicts remain forbidden.
    """

    unresolved = set(unresolved_session_ids)
    if target_session.session_id not in unresolved:
        raise ValueError(
            "Double booking is allowed only for a session left unresolved by the normal plan"
        )

    therapists_tuple = tuple(therapists)
    sessions_tuple = tuple(sessions)
    absences_tuple = tuple(absences)
    patients_tuple = tuple(patients)
    replacements_tuple = tuple(replacements)
    cancellations_tuple = tuple(cancellations)
    base_entries_tuple = tuple(base_entries)

    therapist_by_id = {item.therapist_id: item for item in therapists_tuple}
    therapist = therapist_by_id.get(therapist_id)
    if therapist is None:
        raise ValueError("Unknown double-booking therapist")
    if therapist_id == target_session.therapist_id:
        raise ValueError("Double-booking therapist cannot be the absent original therapist")
    if target_session.robotic and not therapist.robotic_capable:
        raise ValueError("Double-booking therapist is not robotic-capable")

    if any(cancellation.applies_to(target_session) for cancellation in cancellations_tuple):
        raise ValueError("Cancelled session cannot receive a double-booking replacement")
    if any(
        replacement.target_session_id == target_session.session_id
        for replacement in replacements_tuple
    ):
        raise ValueError("Target session already has a replacement")

    if _therapist_absent(
        therapist_id,
        target_session=target_session,
        target_time=replacement_time,
        absences=absences_tuple,
    ):
        raise ValueError("Double-booking therapist is absent at selected time")

    if policy_book is not None:
        validate_policy_replacement_time(
            provider_id=therapist_id,
            target_date=target_session.session_date,
            replacement_time=replacement_time,
            policy_book=policy_book,
        )

    if not is_patient_available(
        patient_id=target_session.patient_id,
        target_date=target_session.session_date,
        target_time=replacement_time,
        sessions=sessions_tuple,
        absences=absences_tuple,
        replacements=replacements_tuple,
        cancellations=cancellations_tuple,
        ignore_session_id=target_session.session_id,
    ):
        raise ValueError("Patient is not available at selected double-booking time")

    validate_patient_cross_specialty_time(
        patient_id=target_session.patient_id,
        target_date=target_session.session_date,
        target_time=replacement_time,
        base_entries=base_entries_tuple,
    )

    if is_therapist_available(
        therapist_id=therapist_id,
        target_date=target_session.session_date,
        target_time=replacement_time,
        sessions=sessions_tuple,
        absences=absences_tuple,
        replacements=replacements_tuple,
        cancellations=cancellations_tuple,
    ):
        raise ValueError(
            "Selected therapist is normally available; double booking is not required"
        )

    if not _has_live_therapist_slot_conflict(
        therapist_id,
        target_session=target_session,
        target_time=replacement_time,
        sessions=sessions_tuple,
        absences=absences_tuple,
        replacements=replacements_tuple,
        cancellations=cancellations_tuple,
    ):
        raise ValueError(
            "Selected therapist is unavailable for a reason other than a live slot conflict"
        )

    return DoubleBookingExceptionProposal(
        target_session_id=target_session.session_id,
        patient_id=target_session.patient_id,
        therapist_id=therapist_id,
        replacement_time=replacement_time,
    )


def confirm_double_booking_exception(
    proposal: DoubleBookingExceptionProposal,
    *,
    target_session: Session,
    replacement_id: str,
    confirmed: bool,
) -> ReplacementAssignment:
    """Materialize the exceptional assignment only after explicit confirmation."""

    if not confirmed:
        raise ValueError("Explicit confirmation is required for double booking")
    if not replacement_id.strip():
        raise ValueError("replacement_id is required")
    if proposal.target_session_id != target_session.session_id:
        raise ValueError("Double-booking proposal does not match target session")
    if proposal.patient_id != target_session.patient_id:
        raise ValueError("Double-booking proposal patient does not match target session")

    return ReplacementAssignment(
        replacement_id=replacement_id,
        target_session_id=target_session.session_id,
        patient_id=target_session.patient_id,
        original_therapist_id=target_session.therapist_id,
        replacement_therapist_id=proposal.therapist_id,
        replacement_date=target_session.session_date,
        replacement_time=proposal.replacement_time,
        reason="CONFIRMED DOUBLE BOOKING EXCEPTION",
        replacement_provider_kind=ReplacementProviderKind.THERAPIST,
    )
