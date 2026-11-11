"""Pure, fail-closed lifecycle rules for the synthetic Diomil pilot.

A lifecycle event limits *future* schedule generation without destroying any
source record or finalized clinical history. Persistence and VBA menus must
implement a separate, audited native-Excel transaction.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from enum import Enum
from typing import Iterable

from .models import BaseScheduleEntry, Patient, PatientType, Session, Student, Therapist


class LifecycleError(ValueError):
    """Unsafe lifecycle transition or missing exact roster identity."""


class PatientExitKind(str, Enum):
    OUTPATIENT_PROGRAM_END = "outpatient_program_end"
    INPATIENT_DISCHARGE = "inpatient_discharge"


@dataclass(frozen=True)
class PatientExit:
    patient_id: str
    kind: PatientExitKind
    last_active_date: date
    reason: str = ""

    def __post_init__(self) -> None:
        if not self.patient_id.strip():
            raise LifecycleError("PatientID is required")

    def permits(self, treatment_date: date) -> bool:
        return treatment_date <= self.last_active_date


@dataclass(frozen=True)
class ProviderRetirement:
    provider_id: str
    last_active_date: date
    is_student: bool = False

    def __post_init__(self) -> None:
        if not self.provider_id.strip():
            raise LifecycleError("Provider ID is required")


@dataclass(frozen=True)
class LifecyclePreview:
    patient_exit: PatientExit
    affected_entries: tuple[str, ...]
    future_session_ids: tuple[str, ...]


def prepare_patient_exit(
    *,
    patient: Patient,
    kind: PatientExitKind,
    last_active_date: date,
    entries: Iterable[BaseScheduleEntry],
    scheduled_sessions: Iterable[Session] = (),
) -> LifecyclePreview:
    """Preview the exact impact, without mutating the program or history."""
    expected = (
        PatientExitKind.OUTPATIENT_PROGRAM_END
        if patient.patient_type == PatientType.OUTPATIENT
        else PatientExitKind.INPATIENT_DISCHARGE
    )
    if kind != expected:
        raise LifecycleError("Exit action does not match the registered patient type")
    event = PatientExit(patient.patient_id, kind, last_active_date)
    entry_ids = tuple(sorted({
        item.base_entry_id for item in entries if item.patient_id == patient.patient_id
    }))
    upcoming = tuple(sorted({
        session.session_id for session in scheduled_sessions
        if session.patient_id == patient.patient_id
        and session.session_date > last_active_date
    }))
    return LifecyclePreview(event, entry_ids, upcoming)


def patient_active_on(
    patient_id: str, target_date: date, exits: Iterable[PatientExit]
) -> bool:
    """Default active until one explicit, nonduplicated exit is recorded."""
    matches = [item for item in exits if item.patient_id == patient_id]
    if len(matches) > 1:
        raise LifecycleError(f"Multiple lifecycle exits for PatientID {patient_id!r}")
    return not matches or matches[0].permits(target_date)


def active_entries_on(
    entries: Iterable[BaseScheduleEntry],
    target_date: date,
    exits: Iterable[PatientExit],
) -> list[BaseScheduleEntry]:
    """Filter recurring entries by patient exit, not by patient absence flags."""
    events = tuple(exits)
    return [
        entry for entry in entries
        if patient_active_on(entry.patient_id, target_date, events)
    ]


def _require_new_id(identity: str, existing: Iterable[str]) -> None:
    if not identity.strip():
        raise LifecycleError("A nonblank unique ID is required")
    if identity.casefold() in {item.casefold() for item in existing}:
        raise LifecycleError(f"Duplicate roster ID: {identity!r}")


def add_therapist(
    therapists: Iterable[Therapist], new_therapist: Therapist
) -> tuple[Therapist, ...]:
    current = tuple(therapists)
    _require_new_id(new_therapist.therapist_id, (x.therapist_id for x in current))
    if not new_therapist.display_name.strip():
        raise LifecycleError("Therapist name is required")
    return (*current, new_therapist)


def add_student(
    students: Iterable[Student], new_student: Student
) -> tuple[Student, ...]:
    current = tuple(students)
    _require_new_id(new_student.student_id, (x.student_id for x in current))
    if new_student.placement_end < new_student.placement_start:
        raise LifecycleError("Student placement end cannot precede start")
    if not new_student.display_name.strip():
        raise LifecycleError("Student name is required")
    return (*current, new_student)


def retire_provider(
    *,
    provider_id: str,
    last_active_date: date,
    entries: Iterable[BaseScheduleEntry],
    is_student: bool = False,
    affected_student_entry_ids: Iterable[str] = (),
    resolved_entry_ids: Iterable[str] = (),
) -> ProviderRetirement:
    """Fail closed if active recurring work has not been reassigned.

    For students, callers supply resolved source entry IDs from the separate
    StudentAssignment -> base-entry mapping. A student must not be treated as a
    regular therapist merely because names are similar.
    """
    if not provider_id.strip():
        raise LifecycleError("Provider ID is required")
    if is_student:
        affected = set(affected_student_entry_ids)
    else:
        affected = {
            entry.base_entry_id for entry in entries
            if entry.therapist_id == provider_id
        }
    unresolved = affected - set(resolved_entry_ids)
    if unresolved:
        raise LifecycleError(
            "Provider owns unresolved recurring assignments: " +
            ", ".join(sorted(unresolved))
        )
    return ProviderRetirement(provider_id, last_active_date, is_student)


def provider_active_on(
    provider_id: str, target_date: date, retirements: Iterable[ProviderRetirement],
) -> bool:
    matches = [x for x in retirements if x.provider_id == provider_id]
    if len(matches) > 1:
        raise LifecycleError(f"Multiple retirement events for {provider_id!r}")
    return not matches or target_date <= matches[0].last_active_date
