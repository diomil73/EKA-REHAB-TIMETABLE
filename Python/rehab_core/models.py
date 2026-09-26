from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from enum import Enum
from typing import Optional


class AbsenceKind(str, Enum):
    PATIENT = "patient"
    THERAPIST = "therapist"
    STUDENT = "student"


class ReplacementProviderKind(str, Enum):
    THERAPIST = "therapist"
    STUDENT = "student"


class PatientType(str, Enum):
    """Operational patient classification.

    Patient type controls visibility/presentation. It must not remove an
    outpatient from workload, capacity, replacement, or productivity logic.
    """

    INPATIENT = "inpatient"
    OUTPATIENT = "outpatient"


class SessionCancellationKind(str, Enum):
    """Why one scheduled treatment did not take place on a specific day."""

    DEPARTMENT_POSTPONED = "department_postponed"
    PATIENT_NO_SHOW = "patient_no_show"


@dataclass(frozen=True)
class Patient:
    patient_id: str
    display_name: str
    room: Optional[str] = None
    infectious: bool = False
    status: Optional[str] = None
    patient_type: PatientType = PatientType.INPATIENT
    hospital_mrn: Optional[str] = None

    def __post_init__(self) -> None:
        # Confirmed business rule: an outpatient is never classified/rendered
        # as infectious. Keep this impossible state out of the domain model,
        # not only out of the registration UI.
        if self.patient_type == PatientType.OUTPATIENT and self.infectious:
            raise ValueError("Outpatient patient cannot be marked infectious")

    @property
    def is_outpatient(self) -> bool:
        return self.patient_type == PatientType.OUTPATIENT


@dataclass(frozen=True)
class Therapist:
    therapist_id: str
    display_name: str
    robotic_capable: bool = False
    # Confirmed operational rule: a physiotherapist may occupy at most six
    # distinct timeslots on any one day. The field lives on the provider model
    # so exceptional staff limits can be represented later without changing the
    # scheduling engine.
    max_daily_timeslots: int = 6


@dataclass(frozen=True)
class BaseScheduleEntry:
    """Recurring schedule row imported from the Excel base programme."""

    base_entry_id: str
    patient_id: str
    treatment: str
    start_time: time
    day_pattern: str
    therapist_id: Optional[str] = None
    robotic: bool = False


@dataclass(frozen=True)
class Session:
    session_id: str
    patient_id: str
    therapist_id: str
    session_date: date
    start_time: time
    treatment: Optional[str] = None
    robotic: bool = False


@dataclass(frozen=True)
class DailySessionCancellation:
    """Cancel one concrete daily session without changing its recurring plan.

    The therapist slot becomes operationally free for other work/replacements,
    while the immutable Session remains available for audit/history.
    """

    cancellation_id: str
    target_session_id: str
    cancellation_date: date
    kind: SessionCancellationKind
    reason: Optional[str] = None

    def applies_to(self, session: Session) -> bool:
        return (
            self.target_session_id == session.session_id
            and self.cancellation_date == session.session_date
        )


@dataclass(frozen=True)
class Student:
    """Student placement kept separate from the Therapist registry.

    Students may receive patients/replacements when they are active and
    replacement-capable. Their confirmed daily capacity is five timeslots.
    """

    student_id: str
    display_name: str
    student_number: int
    placement_start: date
    placement_end: date
    supervisor_therapist_id: Optional[str] = None
    replacement_capable: bool = True
    robotic_capable: bool = False
    max_daily_timeslots: int = 5

    def is_active(self, target_date: date) -> bool:
        return self.placement_start <= target_date <= self.placement_end


@dataclass(frozen=True)
class StudentAssignment:
    """Attach a student to a base session while preserving student identity."""

    assignment_id: str
    student_id: str
    session_id: str


@dataclass(frozen=True)
class DailyAbsence:
    absence_kind: AbsenceKind
    subject_id: str
    absence_date: date
    start_time: Optional[time] = None
    end_time: Optional[time] = None
    reason: Optional[str] = None

    def covers(self, target_date: date, target_time: time) -> bool:
        if target_date != self.absence_date:
            return False
        if self.start_time is None and self.end_time is None:
            return True
        if self.start_time is not None and target_time < self.start_time:
            return False
        if self.end_time is not None and target_time >= self.end_time:
            return False
        return True


@dataclass(frozen=True)
class ReplacementAssignment:
    """Daily overlay for one base session.

    The legacy field name ``replacement_therapist_id`` is intentionally kept
    for compatibility with the existing code. When ``replacement_provider_kind``
    is ``STUDENT``, this field stores the student id.
    """

    replacement_id: str
    target_session_id: str
    patient_id: str
    original_therapist_id: str
    replacement_therapist_id: str
    replacement_date: date
    replacement_time: time
    reason: Optional[str] = None
    replacement_provider_kind: ReplacementProviderKind = ReplacementProviderKind.THERAPIST
