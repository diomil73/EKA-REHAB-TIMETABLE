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
