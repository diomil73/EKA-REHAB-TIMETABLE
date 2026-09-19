from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date, time
from typing import Iterable

from .base_schedule import pattern_applies_on_date
from .day_patterns import RehabWeekday, parse_day_pattern
from .models import BaseScheduleEntry


@dataclass(frozen=True)
class SlotGroupMember:
    """One patient-owned recurring assignment inside a visual timetable slot.

    The member remains independent even when it is rendered together with other
    members in the same therapist/time cell.
    """

    assignment_id: str
    patient_id: str
    therapist_id: str
    start_time: time
    day_pattern: str
    treatment: str
    robotic: bool
    weekdays: frozenset[RehabWeekday]

    @classmethod
    def from_entry(cls, entry: BaseScheduleEntry) -> "SlotGroupMember":
        if entry.therapist_id is None:
            raise ValueError("Cannot build a therapist slot member without therapist_id")
        return cls(
            assignment_id=entry.base_entry_id,
            patient_id=entry.patient_id,
            therapist_id=entry.therapist_id,
            start_time=entry.start_time,
            day_pattern=entry.day_pattern,
            treatment=entry.treatment,
            robotic=entry.robotic,
            weekdays=parse_day_pattern(entry.day_pattern),
        )


@dataclass(frozen=True)
class SlotOverlap:
    first_assignment_id: str
    second_assignment_id: str
    weekdays: frozenset[RehabWeekday]


@dataclass(frozen=True)
class SlotGroup:
    """Derived visual group for one therapist and one clock time.

    A SlotGroup is NOT source data. Its members are independent patient
    assignments. The group is rebuilt whenever therapist/time/day-pattern data
    change.
    """

    therapist_id: str
    start_time: time
    members: tuple[SlotGroupMember, ...]
    overlaps: tuple[SlotOverlap, ...] = ()

    @property
    def is_conflict_free(self) -> bool:
        return not self.overlaps

    @property
    def distinct_patient_ids(self) -> frozenset[str]:
        return frozenset(member.patient_id for member in self.members)

    @property
    def is_same_patient_split(self) -> bool:
        """True when one patient owns multiple complementary assignments here.

        This is not a patient pair. It is one patient whose recurring schedule is
        represented by more than one assignment, for example because different
        day patterns carry different operational attributes.
        """

        return (
            len(self.members) > 1
            and len(self.distinct_patient_ids) == 1
            and self.is_conflict_free
        )

    @property
    def is_pair(self) -> bool:
        """Two distinct patients sharing a conflict-free therapist/time slot."""

        return (
            len(self.members) == 2
            and len(self.distinct_patient_ids) == 2
            and self.is_conflict_free
        )

    def active_members_on(self, target_date: date) -> tuple[SlotGroupMember, ...]:
        """Return only the member(s) that actually have treatment that date."""

        return tuple(
            member
            for member in self.members
            if pattern_applies_on_date(member.day_pattern, target_date)
        )


def _find_overlaps(members: tuple[SlotGroupMember, ...]) -> tuple[SlotOverlap, ...]:
    overlaps: list[SlotOverlap] = []
    for index, first in enumerate(members):
        for second in members[index + 1 :]:
            shared = first.weekdays & second.weekdays
            if shared:
                overlaps.append(
                    SlotOverlap(
                        first_assignment_id=first.assignment_id,
                        second_assignment_id=second.assignment_id,
                        weekdays=frozenset(shared),
                    )
                )
    return tuple(overlaps)


def build_slot_groups(entries: Iterable[BaseScheduleEntry]) -> list[SlotGroup]:
    """Derive timetable slot groups from independent patient assignments.

    Grouping rule:
      same therapist + same time -> same visual slot group.

    Day patterns are deliberately not part of the key. Instead they stay on
    each member and are checked for overlap. This means complementary patterns
    form a clean pair automatically, while accidental same-day collisions are
    still visible and can be reported as conflicts.

    Entries without a therapist remain valid base data but are not included in
    therapist slot groups.
    """

    grouped: dict[tuple[str, time], list[SlotGroupMember]] = {}
    for entry in entries:
        if entry.therapist_id is None:
            continue
        member = SlotGroupMember.from_entry(entry)
        grouped.setdefault((entry.therapist_id, entry.start_time), []).append(member)

    result: list[SlotGroup] = []
    for (therapist_id, start_time), raw_members in grouped.items():
        members = tuple(
            sorted(
                raw_members,
                key=lambda member: (member.patient_id, member.assignment_id),
            )
        )
        result.append(
            SlotGroup(
                therapist_id=therapist_id,
                start_time=start_time,
                members=members,
                overlaps=_find_overlaps(members),
            )
        )

    result.sort(key=lambda group: (group.start_time, group.therapist_id.casefold()))
    return result


def change_patient_assignment(
    entries: Iterable[BaseScheduleEntry],
    *,
    assignment_id: str,
    therapist_id: str | None = None,
    start_time: time | None = None,
    day_pattern: str | None = None,
) -> tuple[list[BaseScheduleEntry], list[SlotGroup]]:
    """Return updated independent assignments and freshly derived slot groups.

    This is intentionally a pure function for planning/tests. It changes only
    the requested patient assignment. Any old pair/group dissolves or reforms
    automatically when ``build_slot_groups`` is run again.

    ``None`` means "leave this field unchanged" in this helper.
    """

    source = list(entries)
    found = False
    updated: list[BaseScheduleEntry] = []

    for entry in source:
        if entry.base_entry_id != assignment_id:
            updated.append(entry)
            continue

        found = True
        replacement = replace(
            entry,
            therapist_id=entry.therapist_id if therapist_id is None else therapist_id,
            start_time=entry.start_time if start_time is None else start_time,
            day_pattern=entry.day_pattern if day_pattern is None else day_pattern,
        )
        # Validate a changed day pattern immediately instead of allowing bad
        # schedule data to travel further into the timetable.
        parse_day_pattern(replacement.day_pattern)
        updated.append(replacement)

    if not found:
        raise KeyError(f"Unknown assignment_id: {assignment_id}")

    return updated, build_slot_groups(updated)
