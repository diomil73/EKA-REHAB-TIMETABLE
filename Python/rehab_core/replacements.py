from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from .availability import (
    is_patient_available,
    is_student_available,
    is_therapist_available,
)
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
from .workload import calculate_student_workload, calculate_therapist_workload


@dataclass(frozen=True)
class ReplacementCandidate:
    """One operational replacement option."""

    therapist_id: str
    display_name: str
    active_sessions: int
    infectious_sessions: int
    robotic_sessions: int
    replacement_sessions: int
    provider_kind: ReplacementProviderKind = ReplacementProviderKind.THERAPIST
    exact_time_available: bool = True
    available_timeslots: tuple[time, ...] = ()
    capacity_limit: int | None = None
    capacity_remaining: int | None = None

    @property
    def provider_id(self) -> str:
        return self.therapist_id


def _day_timeslots(
    target_session: Session,
    sessions: Iterable[Session],
    requested_time: time,
    timeslots: Iterable[time],
) -> tuple[time, ...]:
    explicit = tuple(timeslots)
    if explicit:
        values = set(explicit)
    else:
        values = {
            session.start_time
            for session in sessions
            if session.session_date == target_session.session_date
        }
    values.add(requested_time)
    return tuple(sorted(values))


def _patient_can_use_slot(
    *,
    target_session: Session,
    candidate_time: time,
    sessions: tuple[Session, ...],
    absences: tuple[DailyAbsence, ...],
    replacements: tuple[ReplacementAssignment, ...],
    cancellations: tuple[DailySessionCancellation, ...],
) -> bool:
    return is_patient_available(
        patient_id=target_session.patient_id,
        target_date=target_session.session_date,
        target_time=candidate_time,
        sessions=sessions,
        absences=absences,
        replacements=replacements,
        cancellations=cancellations,
        ignore_session_id=target_session.session_id,
    )


def find_replacement_candidates(
    target_session: Session,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    *,
    replacement_time: time | None = None,
    timeslots: Iterable[time] = (),
    students: Iterable[Student] = (),
    student_assignments: Iterable[StudentAssignment] = (),
) -> list[ReplacementCandidate]:
    """Rank replacement providers using the confirmed operational priority.

    Daily cancellations remove only the concrete cancelled occurrence from
    provider/patient occupancy and workload. The recurring Session stays intact.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    patients = tuple(patients)
    replacements = tuple(replacements)
    cancellations = tuple(cancellations)
    therapists = tuple(therapists)
    students = tuple(students)
    student_assignments = tuple(student_assignments)

    requested_time = replacement_time or target_session.start_time
    day_timeslots = _day_timeslots(
        target_session, sessions, requested_time, timeslots
    )
    patient_by_id = {patient.patient_id: patient for patient in patients}
    target_patient = patient_by_id.get(target_session.patient_id)
    target_is_infectious = bool(target_patient and target_patient.infectious)

    patient_usable_slots = tuple(
        slot
        for slot in day_timeslots
        if _patient_can_use_slot(
            target_session=target_session,
            candidate_time=slot,
            sessions=sessions,
            absences=absences,
            replacements=replacements,
            cancellations=cancellations,
        )
    )
    if not patient_usable_slots:
        return []

    candidates: list[ReplacementCandidate] = []

    for therapist in therapists:
        if therapist.therapist_id == target_session.therapist_id:
            continue
        if target_session.robotic and not therapist.robotic_capable:
            continue

        available_slots = tuple(
            slot
            for slot in patient_usable_slots
            if is_therapist_available(
                therapist_id=therapist.therapist_id,
                target_date=target_session.session_date,
                target_time=slot,
                sessions=sessions,
                absences=absences,
                replacements=replacements,
                cancellations=cancellations,
            )
        )
        if not available_slots:
            continue

        workload = calculate_therapist_workload(
            therapist_id=therapist.therapist_id,
            target_date=target_session.session_date,
            sessions=sessions,
            absences=absences,
            patients=patients,
            replacements=replacements,
            cancellations=cancellations,
        )
        if workload.active_timeslots >= therapist.max_daily_timeslots:
            continue

        candidates.append(
            ReplacementCandidate(
                therapist_id=therapist.therapist_id,
                display_name=therapist.display_name,
                provider_kind=ReplacementProviderKind.THERAPIST,
                active_sessions=workload.active_sessions,
                infectious_sessions=workload.infectious_sessions,
                robotic_sessions=workload.robotic_sessions,
                replacement_sessions=workload.replacement_sessions,
                exact_time_available=requested_time in available_slots,
                available_timeslots=available_slots,
                capacity_limit=therapist.max_daily_timeslots,
                capacity_remaining=therapist.max_daily_timeslots
                - workload.active_timeslots,
            )
        )

    for student in students:
        if not student.is_active(target_session.session_date):
            continue
        if not student.replacement_capable:
            continue
        if target_session.robotic and not student.robotic_capable:
            continue

        workload = calculate_student_workload(
            student_id=student.student_id,
            target_date=target_session.session_date,
            sessions=sessions,
            student_assignments=student_assignments,
            absences=absences,
            patients=patients,
            replacements=replacements,
            cancellations=cancellations,
        )
        if workload.active_timeslots >= student.max_daily_timeslots:
            continue

        available_slots = tuple(
            slot
            for slot in patient_usable_slots
            if is_student_available(
                student_id=student.student_id,
                target_date=target_session.session_date,
                target_time=slot,
                sessions=sessions,
                student_assignments=student_assignments,
                absences=absences,
                replacements=replacements,
                cancellations=cancellations,
            )
        )
        if not available_slots:
            continue

        candidates.append(
            ReplacementCandidate(
                therapist_id=student.student_id,
                display_name=student.display_name,
                provider_kind=ReplacementProviderKind.STUDENT,
                active_sessions=workload.active_sessions,
                infectious_sessions=workload.infectious_sessions,
                robotic_sessions=workload.robotic_sessions,
                replacement_sessions=workload.replacement_sessions,
                exact_time_available=requested_time in available_slots,
                available_timeslots=available_slots,
                capacity_limit=student.max_daily_timeslots,
                capacity_remaining=student.max_daily_timeslots
                - workload.active_timeslots,
            )
        )

    candidates.sort(
        key=lambda candidate: (
            candidate.active_sessions,
            0 if candidate.exact_time_available else 1,
            candidate.infectious_sessions if target_is_infectious else 0,
            candidate.replacement_sessions,
            candidate.display_name.casefold(),
            candidate.therapist_id,
        )
    )
    return candidates


def create_replacement_assignment(
    *,
    replacement_id: str,
    target_session: Session,
    replacement_therapist_id: str,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    replacement_time: time | None = None,
    reason: str | None = None,
    students: Iterable[Student] = (),
    student_assignments: Iterable[StudentAssignment] = (),
    patients: Iterable[Patient] = (),
) -> ReplacementAssignment:
    """Validate and create a daily replacement overlay for therapist/student."""

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)
    cancellations = tuple(cancellations)
    therapists = tuple(therapists)
    students = tuple(students)
    student_assignments = tuple(student_assignments)
    patients = tuple(patients)
    target_time = replacement_time or target_session.start_time

    if any(cancellation.applies_to(target_session) for cancellation in cancellations):
        raise ValueError("Cancelled session cannot receive a replacement")

    if any(
        replacement.target_session_id == target_session.session_id
        for replacement in replacements
    ):
        raise ValueError("Target session already has a replacement")

    therapist_by_id = {
        therapist.therapist_id: therapist for therapist in therapists
    }
    student_by_id = {student.student_id: student for student in students}

    therapist = therapist_by_id.get(replacement_therapist_id)
    student = student_by_id.get(replacement_therapist_id)
    if therapist is not None and student is not None:
        raise ValueError("Provider id exists as both therapist and student")
    if therapist is None and student is None:
        raise ValueError("Unknown replacement provider")

    if therapist is not None:
        if replacement_therapist_id == target_session.therapist_id:
            raise ValueError("Replacement therapist cannot be the original therapist")
        if target_session.robotic and not therapist.robotic_capable:
            raise ValueError("Replacement therapist is not robotic-capable")
        therapist_workload = calculate_therapist_workload(
            therapist_id=therapist.therapist_id,
            target_date=target_session.session_date,
            sessions=sessions,
            absences=absences,
            patients=patients,
            replacements=replacements,
            cancellations=cancellations,
        )
        if therapist_workload.active_timeslots >= therapist.max_daily_timeslots:
            raise ValueError("Therapist has reached daily timeslot capacity")
        provider_kind = ReplacementProviderKind.THERAPIST
    else:
        assert student is not None
        if not student.is_active(target_session.session_date):
            raise ValueError("Student placement is not active")
        if not student.replacement_capable:
            raise ValueError("Student is not replacement-capable")
        if target_session.robotic and not student.robotic_capable:
            raise ValueError("Replacement student is not robotic-capable")
        student_workload = calculate_student_workload(
            student_id=student.student_id,
            target_date=target_session.session_date,
            sessions=sessions,
            student_assignments=student_assignments,
            absences=absences,
            patients=patients,
            replacements=replacements,
            cancellations=cancellations,
        )
        if student_workload.active_timeslots >= student.max_daily_timeslots:
            raise ValueError("Student has reached daily timeslot capacity")
        provider_kind = ReplacementProviderKind.STUDENT

    if not is_patient_available(
        patient_id=target_session.patient_id,
        target_date=target_session.session_date,
        target_time=target_time,
        sessions=sessions,
        absences=absences,
        replacements=replacements,
        cancellations=cancellations,
        ignore_session_id=target_session.session_id,
    ):
        raise ValueError("Patient is not available")

    if provider_kind == ReplacementProviderKind.THERAPIST:
        provider_available = is_therapist_available(
            therapist_id=replacement_therapist_id,
            target_date=target_session.session_date,
            target_time=target_time,
            sessions=sessions,
            absences=absences,
            replacements=replacements,
            cancellations=cancellations,
        )
    else:
        provider_available = is_student_available(
            student_id=replacement_therapist_id,
            target_date=target_session.session_date,
            target_time=target_time,
            sessions=sessions,
            student_assignments=student_assignments,
            absences=absences,
            replacements=replacements,
            cancellations=cancellations,
        )

    if not provider_available:
        raise ValueError("Replacement provider is not available at selected time")

    return ReplacementAssignment(
        replacement_id=replacement_id,
        target_session_id=target_session.session_id,
        patient_id=target_session.patient_id,
        original_therapist_id=target_session.therapist_id,
        replacement_therapist_id=replacement_therapist_id,
        replacement_date=target_session.session_date,
        replacement_time=target_time,
        reason=reason,
        replacement_provider_kind=provider_kind,
    )
