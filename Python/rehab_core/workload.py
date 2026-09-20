from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Iterable

from .models import (
    AbsenceKind,
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
    StudentAssignment,
)


@dataclass(frozen=True)
class TherapistWorkload:
    therapist_id: str
    workload_date: date
    active_sessions: int = 0
    active_timeslots: int = 0
    infectious_sessions: int = 0
    robotic_sessions: int = 0
    replacement_sessions: int = 0


@dataclass(frozen=True)
class StudentWorkload:
    student_id: str
    workload_date: date
    active_sessions: int = 0
    active_timeslots: int = 0
    infectious_sessions: int = 0
    robotic_sessions: int = 0
    replacement_sessions: int = 0


def _patient_absent(
    patient_id: str,
    target_date: date,
    target_time: time,
    absences: Iterable[DailyAbsence],
) -> bool:
    return any(
        absence.absence_kind == AbsenceKind.PATIENT
        and absence.subject_id == patient_id
        and absence.covers(target_date, target_time)
        for absence in absences
    )


def calculate_therapist_workload(
    therapist_id: str,
    target_date: date,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> TherapistWorkload:
    """Calculate the real operational workload for one therapist/day.

    ``active_sessions`` remains a patient/session count because it is useful for
    load balancing. ``active_timeslots`` is deliberately separate and is the
    authoritative number used for the hard capacity rule (max 6 for a normal
    physiotherapist). Two patients that genuinely share one clock slot count as
    two sessions but only one occupied timeslot.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)
    patient_by_id = {patient.patient_id: patient for patient in patients}
    session_by_id = {session.session_id: session for session in sessions}
    replaced_session_ids = {
        replacement.target_session_id for replacement in replacements
    }

    active_sessions = 0
    occupied_times: set[time] = set()
    infectious_sessions = 0
    robotic_sessions = 0
    replacement_sessions = 0

    for session in sessions:
        if session.therapist_id != therapist_id or session.session_date != target_date:
            continue
        if session.session_id in replaced_session_ids:
            continue
        if _patient_absent(
            session.patient_id,
            session.session_date,
            session.start_time,
            absences,
        ):
            continue

        active_sessions += 1
        occupied_times.add(session.start_time)
        patient = patient_by_id.get(session.patient_id)
        if patient is not None and patient.infectious:
            infectious_sessions += 1
        if session.robotic:
            robotic_sessions += 1

    for replacement in replacements:
        if replacement.replacement_provider_kind != ReplacementProviderKind.THERAPIST:
            continue
        if (
            replacement.replacement_therapist_id != therapist_id
            or replacement.replacement_date != target_date
        ):
            continue
        if _patient_absent(
            replacement.patient_id,
            replacement.replacement_date,
            replacement.replacement_time,
            absences,
        ):
            continue

        active_sessions += 1
        occupied_times.add(replacement.replacement_time)
        replacement_sessions += 1
        patient = patient_by_id.get(replacement.patient_id)
        if patient is not None and patient.infectious:
            infectious_sessions += 1

        base_session = session_by_id.get(replacement.target_session_id)
        if base_session is not None and base_session.robotic:
            robotic_sessions += 1

    return TherapistWorkload(
        therapist_id=therapist_id,
        workload_date=target_date,
        active_sessions=active_sessions,
        active_timeslots=len(occupied_times),
        infectious_sessions=infectious_sessions,
        robotic_sessions=robotic_sessions,
        replacement_sessions=replacement_sessions,
    )


def calculate_student_workload(
    student_id: str,
    target_date: date,
    sessions: Iterable[Session],
    student_assignments: Iterable[StudentAssignment] = (),
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> StudentWorkload:
    """Calculate patient/session load and occupied timeslots for a student."""

    sessions = tuple(sessions)
    student_assignments = tuple(student_assignments)
    absences = tuple(absences)
    replacements = tuple(replacements)
    patient_by_id = {patient.patient_id: patient for patient in patients}
    session_by_id = {session.session_id: session for session in sessions}

    active_sessions = 0
    occupied_times: set[time] = set()
    infectious_sessions = 0
    robotic_sessions = 0
    replacement_sessions = 0

    for assignment in student_assignments:
        if assignment.student_id != student_id:
            continue
        session = session_by_id.get(assignment.session_id)
        if session is None or session.session_date != target_date:
            continue
        if _patient_absent(
            session.patient_id,
            session.session_date,
            session.start_time,
            absences,
        ):
            continue

        active_sessions += 1
        occupied_times.add(session.start_time)
        patient = patient_by_id.get(session.patient_id)
        if patient is not None and patient.infectious:
            infectious_sessions += 1
        if session.robotic:
            robotic_sessions += 1

    for replacement in replacements:
        if replacement.replacement_provider_kind != ReplacementProviderKind.STUDENT:
            continue
        if (
            replacement.replacement_therapist_id != student_id
            or replacement.replacement_date != target_date
        ):
            continue
        if _patient_absent(
            replacement.patient_id,
            replacement.replacement_date,
            replacement.replacement_time,
            absences,
        ):
            continue

        active_sessions += 1
        occupied_times.add(replacement.replacement_time)
        replacement_sessions += 1
        patient = patient_by_id.get(replacement.patient_id)
        if patient is not None and patient.infectious:
            infectious_sessions += 1
        base_session = session_by_id.get(replacement.target_session_id)
        if base_session is not None and base_session.robotic:
            robotic_sessions += 1

    return StudentWorkload(
        student_id=student_id,
        workload_date=target_date,
        active_sessions=active_sessions,
        active_timeslots=len(occupied_times),
        infectious_sessions=infectious_sessions,
        robotic_sessions=robotic_sessions,
        replacement_sessions=replacement_sessions,
    )
