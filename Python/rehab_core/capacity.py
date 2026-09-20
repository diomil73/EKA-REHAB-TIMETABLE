from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Iterable

from .day_patterns import RehabWeekday, parse_day_pattern
from .models import (
    BaseScheduleEntry,
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    Session,
    Student,
    StudentAssignment,
    Therapist,
)
from .workload import calculate_student_workload, calculate_therapist_workload


@dataclass(frozen=True)
class DailyCapacityStatus:
    provider_id: str
    capacity_date: date
    used_timeslots: int
    max_timeslots: int

    @property
    def remaining_timeslots(self) -> int:
        return max(self.max_timeslots - self.used_timeslots, 0)

    @property
    def at_capacity(self) -> bool:
        return self.used_timeslots >= self.max_timeslots


@dataclass(frozen=True)
class RecurringCapacityIssue:
    weekday: RehabWeekday
    used_timeslots_before: int
    used_timeslots_after: int
    max_timeslots: int


@dataclass(frozen=True)
class RecurringCapacityCheck:
    provider_id: str
    max_timeslots: int
    issues: tuple[RecurringCapacityIssue, ...]

    @property
    def allowed(self) -> bool:
        return not self.issues


def therapist_daily_capacity(
    therapist: Therapist,
    target_date: date,
    sessions: Iterable[Session],
    *,
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> DailyCapacityStatus:
    workload = calculate_therapist_workload(
        therapist_id=therapist.therapist_id,
        target_date=target_date,
        sessions=sessions,
        absences=absences,
        patients=patients,
        replacements=replacements,
    )
    return DailyCapacityStatus(
        provider_id=therapist.therapist_id,
        capacity_date=target_date,
        used_timeslots=workload.active_timeslots,
        max_timeslots=therapist.max_daily_timeslots,
    )


def student_daily_capacity(
    student: Student,
    target_date: date,
    sessions: Iterable[Session],
    *,
    student_assignments: Iterable[StudentAssignment] = (),
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> DailyCapacityStatus:
    workload = calculate_student_workload(
        student_id=student.student_id,
        target_date=target_date,
        sessions=sessions,
        student_assignments=student_assignments,
        absences=absences,
        patients=patients,
        replacements=replacements,
    )
    return DailyCapacityStatus(
        provider_id=student.student_id,
        capacity_date=target_date,
        used_timeslots=workload.active_timeslots,
        max_timeslots=student.max_daily_timeslots,
    )


def validate_recurring_timeslot_capacity(
    *,
    provider_id: str,
    max_daily_timeslots: int,
    existing_entries: Iterable[BaseScheduleEntry],
    proposed_day_pattern: str,
    proposed_time: time,
) -> RecurringCapacityCheck:
    """Validate a future recurring assignment before it enters the base timetable.

    This is the common capacity primitive for new patients, permanent therapist
    changes and permanent time changes. Capacity is counted as DISTINCT clock
    times per weekday, not number of patient rows.

    Conflict validation remains separate: if a provider already has a patient at
    the same time on the same weekday, the capacity count does not increase but
    the normal conflict engine must still reject the double booking.
    """

    if max_daily_timeslots < 1:
        raise ValueError("max_daily_timeslots must be >= 1")

    occupied: dict[RehabWeekday, set[time]] = {
        weekday: set() for weekday in RehabWeekday
    }
    for entry in existing_entries:
        if entry.therapist_id != provider_id:
            continue
        for weekday in parse_day_pattern(entry.day_pattern):
            occupied[weekday].add(entry.start_time)

    issues: list[RecurringCapacityIssue] = []
    for weekday in parse_day_pattern(proposed_day_pattern):
        before = len(occupied[weekday])
        after = len(occupied[weekday] | {proposed_time})
        if after > max_daily_timeslots:
            issues.append(
                RecurringCapacityIssue(
                    weekday=weekday,
                    used_timeslots_before=before,
                    used_timeslots_after=after,
                    max_timeslots=max_daily_timeslots,
                )
            )

    return RecurringCapacityCheck(
        provider_id=provider_id,
        max_timeslots=max_daily_timeslots,
        issues=tuple(issues),
    )


def validate_new_patient_assignment_capacity(
    *,
    therapist: Therapist,
    existing_entries: Iterable[BaseScheduleEntry],
    day_pattern: str,
    start_time: time,
) -> RecurringCapacityCheck:
    """Named wrapper used by the future new-patient placement workflow."""

    return validate_recurring_timeslot_capacity(
        provider_id=therapist.therapist_id,
        max_daily_timeslots=therapist.max_daily_timeslots,
        existing_entries=existing_entries,
        proposed_day_pattern=day_pattern,
        proposed_time=start_time,
    )
