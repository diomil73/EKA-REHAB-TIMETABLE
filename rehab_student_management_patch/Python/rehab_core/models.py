from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from enum import Enum
from typing import Optional


class AbsenceKind(str, Enum):
    PATIENT = "patient"
    THERAPIST = "therapist"


@dataclass(frozen=True)
class Patient:
    patient_id: str
    display_name: str
    room: Optional[str] = None
    infectious: bool = False
    status: Optional[str] = None


@dataclass(frozen=True)
class Therapist:
    therapist_id: str
    display_name: str
    robotic_capable: bool = False


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
class Student:
    """A student placement kept separate from the Therapist registry.

    ``student_number`` is a stable internal presentation number used only after
    the placement has ended. The real name is always retained internally.
    """

    student_id: str
    display_name: str
    student_number: int
    placement_start: date
    placement_end: date
    supervisor_therapist_id: Optional[str] = None

    def is_active(self, target_date: date) -> bool:
        return self.placement_start <= target_date <= self.placement_end


@dataclass(frozen=True)
class StudentAssignment:
    """Attach a student to a base session without making them a therapist."""

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
    """A daily overlay that replaces the therapist for one base session.

    The base Session is deliberately not edited. This object represents the
    operational change for a single day and time.
    """

    replacement_id: str
    target_session_id: str
    patient_id: str
    original_therapist_id: str
    replacement_therapist_id: str
    replacement_date: date
    replacement_time: time
    reason: Optional[str] = None
