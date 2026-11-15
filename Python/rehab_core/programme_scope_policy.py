"""Read-only scheduling policy for morning/afternoon care and medical deferrals.

Pure contract: no Excel, VBA, persistence, patient names or patient migration.
Operational UI integration requires separate tests and native Excel acceptance.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from enum import Enum
from typing import Iterable


class SchedulingScopeError(ValueError):
    """The requested cancellation scope is inconsistent or ambiguous."""


class Programme(str, Enum):
    MORNING = "MORNING"
    AFTERNOON = "AFTERNOON"


class AbsenceReason(str, Enum):
    MEDICAL = "MEDICAL"
    TEMPORARY_DISCHARGE = "TEMPORARY_DISCHARGE"
    TRANSFER = "TRANSFER"
    SESSION_DEFERRAL = "SESSION_DEFERRAL"
    THERAPIST_ABSENCE = "THERAPIST_ABSENCE"


class AbsenceScope(str, Enum):
    FULL_DAY = "FULL_DAY"
    TIME_WINDOW = "TIME_WINDOW"
    SESSION = "SESSION"


@dataclass(frozen=True)
class PlannedSession:
    session_id: str
    patient_id: str
    day: date
    programme: Programme
    treatment: str
    start: time
    end: time
    therapist_id: str | None

    def __post_init__(self) -> None:
        if not self.session_id or not self.patient_id or not self.treatment:
            raise SchedulingScopeError("Session ID, patient ID and treatment required")
        if not isinstance(self.programme, Programme):
            raise SchedulingScopeError("Unknown programme")
        if not isinstance(self.day, date) or isinstance(self.day, datetime):
            raise SchedulingScopeError("Explicit session date required")
        if not isinstance(self.start, time) or not isinstance(self.end, time) or self.start >= self.end:
            raise SchedulingScopeError("Positive within-day interval required")
        if self.programme is Programme.AFTERNOON:
            if self.treatment != "ΦΘ" or not self.therapist_id:
                raise SchedulingScopeError("Afternoon permits named physiotherapy only")


@dataclass(frozen=True)
class Absence:
    patient_id: str
    day: date
    reason: AbsenceReason
    scope: AbsenceScope
    programme: Programme | None = None
    start: time | None = None
    end: time | None = None
    session_id: str | None = None

    def __post_init__(self) -> None:
        if not self.patient_id or not isinstance(self.day, date) or isinstance(self.day, datetime):
            raise SchedulingScopeError("Patient and exact date required")
        if not isinstance(self.reason, AbsenceReason) or not isinstance(self.scope, AbsenceScope):
            raise SchedulingScopeError("Known reason and explicit scope required")
        if self.programme is not None and not isinstance(self.programme, Programme):
            raise SchedulingScopeError("Unknown programme")
        if self.reason in (AbsenceReason.TEMPORARY_DISCHARGE, AbsenceReason.TRANSFER):
            if self.scope is not AbsenceScope.FULL_DAY:
                raise SchedulingScopeError("Transfer/discharge are full-day states")
        if self.scope is AbsenceScope.FULL_DAY:
            if self.programme is not None or any(x is not None for x in (self.start, self.end, self.session_id)):
                raise SchedulingScopeError("Full-day absence cannot narrow its coverage")
            if self.reason not in (
                AbsenceReason.MEDICAL, AbsenceReason.TEMPORARY_DISCHARGE, AbsenceReason.TRANSFER,
            ):
                raise SchedulingScopeError("Therapist/session deferrals must be targeted")
        elif self.scope is AbsenceScope.TIME_WINDOW:
            if (
                not isinstance(self.start, time) or not isinstance(self.end, time)
                or self.start >= self.end or self.session_id is not None
            ):
                raise SchedulingScopeError("Exact nonempty time interval required")
            if self.reason is not AbsenceReason.MEDICAL:
                raise SchedulingScopeError("Only medical deferral uses a patient time interval")
        elif self.scope is AbsenceScope.SESSION:
            if not self.session_id or self.start is not None or self.end is not None:
                raise SchedulingScopeError("Exact target session required")


def affected_session_ids(sessions: Iterable[PlannedSession], absence: Absence) -> tuple[str, ...]:
    """Preview targets; never cancel or change the recurring weekly programme.

    FULL_DAY covers all sessions, morning and afternoon. TIME_WINDOW uses real
    interval overlap. SESSION targets exactly one session, even if the patient
    has another session at the same clock time in the other programme.
    """
    rows = tuple(sessions)
    if len({s.session_id for s in rows}) != len(rows):
        raise SchedulingScopeError("Nonunique session IDs")
    same_day = tuple(s for s in rows if s.patient_id == absence.patient_id and s.day == absence.day)
    if absence.scope is AbsenceScope.FULL_DAY:
        return tuple(s.session_id for s in same_day)
    if absence.scope is AbsenceScope.TIME_WINDOW:
        return tuple(
            s.session_id for s in same_day
            if (absence.programme is None or s.programme is absence.programme)
            and s.start < absence.end and absence.start < s.end
        )
    target = next((s for s in same_day if s.session_id == absence.session_id), None)
    if target is None or (absence.programme is not None and target.programme is not absence.programme):
        raise SchedulingScopeError("Target does not match patient/date/programme")
    return (target.session_id,)


def conflicting_session_pairs(sessions: Iterable[PlannedSession]) -> tuple[tuple[str, str], ...]:
    """Conflict checks are completely independent per programme.

    Within each programme, two sessions conflict if they overlap and share
    the same patient or the same nonempty therapist. The caller must provide
    actual start/end intervals; we never invent treatment durations.
    """
    rows = tuple(sessions)
    if len({s.session_id for s in rows}) != len(rows):
        raise SchedulingScopeError("Nonunique session IDs")
    conflicts = []
    for index, a in enumerate(rows):
        for b in rows[index + 1:]:
            if a.day != b.day or a.programme is not b.programme:
                continue
            if a.start >= b.end or b.start >= a.end:
                continue
            if a.patient_id == b.patient_id or (
                a.therapist_id and b.therapist_id and a.therapist_id == b.therapist_id
            ):
                conflicts.append((a.session_id, b.session_id))
    return tuple(conflicts)
