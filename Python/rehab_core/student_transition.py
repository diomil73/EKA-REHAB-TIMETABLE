from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time, timedelta
import re
from typing import Iterable, Mapping

from .models import BaseScheduleEntry, Session, Student, StudentAssignment
from .permanent_assignment import PermanentAssignmentCheck, check_permanent_assignment


_DATED_SESSION_ID = re.compile(r"^(?P<base>.+)@\d{4}-\d{2}-\d{2}$")


@dataclass(frozen=True)
class StudentEndAffectedEntry:
    """Recurring source entry currently associated with one student.

    The assignment ids are kept for audit. No workbook/base-schedule write is
    performed by this module.
    """

    student_id: str
    source_entry: BaseScheduleEntry
    assignment_ids: tuple[str, ...]

    @property
    def patient_id(self) -> str:
        return self.source_entry.patient_id


@dataclass(frozen=True)
class StudentEndReassignmentOption:
    affected_entry: StudentEndAffectedEntry
    destination_provider_id: str
    destination_time: time
    destination_check: PermanentAssignmentCheck

    @property
    def same_time(self) -> bool:
        return self.destination_time == self.affected_entry.source_entry.start_time

    @property
    def destination_peak_after(self) -> int:
        return max(
            (projection.used_timeslots_after for projection in self.destination_check.days),
            default=0,
        )

    @property
    def destination_total_after(self) -> int:
        return sum(
            projection.used_timeslots_after for projection in self.destination_check.days
        )


@dataclass(frozen=True)
class StudentEndTransitionPlan:
    student_id: str
    student_display_name: str
    expired_label: str
    placement_end: date
    effective_from: date
    affected_entries: tuple[StudentEndAffectedEntry, ...]
    options: tuple[StudentEndReassignmentOption, ...]

    @property
    def affected_patient_ids(self) -> tuple[str, ...]:
        return tuple(dict.fromkeys(item.patient_id for item in self.affected_entries))

    def options_for_entry(self, base_entry_id: str) -> tuple[StudentEndReassignmentOption, ...]:
        return tuple(
            option
            for option in self.options
            if option.affected_entry.source_entry.base_entry_id == base_entry_id
        )


def expired_student_label(student: Student) -> str:
    """Stable post-placement display label while preserving real identity internally."""

    return f"Φοιτ.{student.student_number}"


def _base_entry_id_from_session_id(session_id: str) -> str | None:
    match = _DATED_SESSION_ID.match(session_id)
    if match:
        return match.group("base")
    return session_id or None


def resolve_student_base_entries(
    *,
    student_id: str,
    student_assignments: Iterable[StudentAssignment],
    sessions: Iterable[Session],
    base_entries: Iterable[BaseScheduleEntry],
) -> tuple[StudentEndAffectedEntry, ...]:
    """Resolve a student's operational assignments back to recurring base entries.

    Materialized session ids use ``<base_entry_id>@YYYY-MM-DD``. Exact base-entry
    ids are also accepted so callers can use persistent assignment records later.
    Unresolvable assignments are intentionally ignored rather than guessed.
    """

    entries_by_id = {entry.base_entry_id: entry for entry in base_entries}
    sessions_by_id = {session.session_id: session for session in sessions}
    assignment_ids_by_entry: dict[str, list[str]] = {}

    for assignment in student_assignments:
        if assignment.student_id != student_id:
            continue

        session = sessions_by_id.get(assignment.session_id)
        session_id = session.session_id if session is not None else assignment.session_id
        base_entry_id = _base_entry_id_from_session_id(session_id)
        if base_entry_id not in entries_by_id:
            continue
        assignment_ids_by_entry.setdefault(base_entry_id, []).append(assignment.assignment_id)

    affected = [
        StudentEndAffectedEntry(
            student_id=student_id,
            source_entry=entries_by_id[base_entry_id],
            assignment_ids=tuple(dict.fromkeys(assignment_ids)),
        )
        for base_entry_id, assignment_ids in assignment_ids_by_entry.items()
    ]
    affected.sort(
        key=lambda item: (
            item.source_entry.patient_id.casefold(),
            item.source_entry.start_time,
            item.source_entry.base_entry_id,
        )
    )
    return tuple(affected)


def plan_student_end_transition(
    *,
    student: Student,
    student_assignments: Iterable[StudentAssignment],
    sessions: Iterable[Session],
    base_entries: Iterable[BaseScheduleEntry],
    destination_provider_ids: Iterable[str],
    standard_timeslots: Iterable[time],
    provider_capacity_limits: Mapping[str, int] | None = None,
) -> StudentEndTransitionPlan:
    """Build read-only permanent reassignment suggestions for an ending placement.

    The planner never edits the base schedule or student assignments. Every
    suggestion passes the existing permanent weekly provider-capacity,
    provider-collision, and cross-specialty patient-conflict gate.
    """

    entries = tuple(base_entries)
    affected = resolve_student_base_entries(
        student_id=student.student_id,
        student_assignments=student_assignments,
        sessions=sessions,
        base_entries=entries,
    )

    destinations = tuple(dict.fromkeys(provider_id for provider_id in destination_provider_ids if provider_id))
    timeslots = tuple(dict.fromkeys(standard_timeslots))
    limits = dict(provider_capacity_limits or {})

    options: list[StudentEndReassignmentOption] = []
    for item in affected:
        source = item.source_entry
        for provider_id in destinations:
            max_slots = limits.get(provider_id, 6)
            for destination_time in timeslots:
                check = check_permanent_assignment(
                    provider_id=provider_id,
                    max_daily_timeslots=max_slots,
                    existing_entries=entries,
                    proposed_day_pattern=source.day_pattern,
                    proposed_time=destination_time,
                    source_entry_id=source.base_entry_id,
                    patient_id=source.patient_id,
                )
                if not check.allowed:
                    continue
                options.append(
                    StudentEndReassignmentOption(
                        affected_entry=item,
                        destination_provider_id=provider_id,
                        destination_time=destination_time,
                        destination_check=check,
                    )
                )

    options.sort(
        key=lambda option: (
            option.affected_entry.source_entry.patient_id.casefold(),
            option.affected_entry.source_entry.base_entry_id,
            0 if option.same_time else 1,
            option.destination_peak_after,
            option.destination_total_after,
            option.destination_provider_id.casefold(),
            option.destination_time,
        )
    )

    return StudentEndTransitionPlan(
        student_id=student.student_id,
        student_display_name=student.display_name,
        expired_label=expired_student_label(student),
        placement_end=student.placement_end,
        effective_from=student.placement_end + timedelta(days=1),
        affected_entries=affected,
        options=tuple(options),
    )
