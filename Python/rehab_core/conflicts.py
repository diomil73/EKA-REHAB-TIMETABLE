from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from enum import Enum
from typing import Iterable

from .models import (
    AbsenceKind,
    DailyAbsence,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
)


class ConflictKind(str, Enum):
    THERAPIST_DOUBLE_BOOKING = "therapist_double_booking"
    STUDENT_DOUBLE_BOOKING = "student_double_booking"
    PATIENT_DOUBLE_BOOKING = "patient_double_booking"


@dataclass(frozen=True)
class OperationalOccurrence:
    source_id: str
    patient_id: str
    therapist_id: str
    occurrence_date: date
    start_time: time
    is_replacement: bool = False
    provider_kind: ReplacementProviderKind = ReplacementProviderKind.THERAPIST

    @property
    def provider_id(self) -> str:
        return self.therapist_id


@dataclass(frozen=True)
class ScheduleConflict:
    kind: ConflictKind
    subject_id: str
    conflict_date: date
    start_time: time
    occurrence_ids: tuple[str, ...]


def _is_patient_absent(
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


def build_operational_occurrences(
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> list[OperationalOccurrence]:
    """Build today's effective schedule without mutating the base schedule."""

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)
    replacement_by_session = {
        replacement.target_session_id: replacement for replacement in replacements
    }

    occurrences: list[OperationalOccurrence] = []

    for session in sessions:
        replacement = replacement_by_session.get(session.session_id)
        if replacement is not None:
            continue
        if _is_patient_absent(
            session.patient_id,
            session.session_date,
            session.start_time,
            absences,
        ):
            continue
        occurrences.append(
            OperationalOccurrence(
                source_id=session.session_id,
                patient_id=session.patient_id,
                therapist_id=session.therapist_id,
                occurrence_date=session.session_date,
                start_time=session.start_time,
                provider_kind=ReplacementProviderKind.THERAPIST,
            )
        )

    for replacement in replacements:
        if _is_patient_absent(
            replacement.patient_id,
            replacement.replacement_date,
            replacement.replacement_time,
            absences,
        ):
            continue
        occurrences.append(
            OperationalOccurrence(
                source_id=replacement.replacement_id,
                patient_id=replacement.patient_id,
                therapist_id=replacement.replacement_therapist_id,
                occurrence_date=replacement.replacement_date,
                start_time=replacement.replacement_time,
                is_replacement=True,
                provider_kind=replacement.replacement_provider_kind,
            )
        )

    return occurrences


def find_operational_conflicts(
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
) -> list[ScheduleConflict]:
    """Find provider and patient double bookings in the effective day."""

    occurrences = build_operational_occurrences(sessions, absences, replacements)
    conflicts: list[ScheduleConflict] = []

    provider_slots: dict[
        tuple[ReplacementProviderKind, str, date, time], list[OperationalOccurrence]
    ] = {}
    patient_slots: dict[tuple[str, date, time], list[OperationalOccurrence]] = {}

    for occurrence in occurrences:
        provider_slots.setdefault(
            (
                occurrence.provider_kind,
                occurrence.provider_id,
                occurrence.occurrence_date,
                occurrence.start_time,
            ),
            [],
        ).append(occurrence)
        patient_slots.setdefault(
            (
                occurrence.patient_id,
                occurrence.occurrence_date,
                occurrence.start_time,
            ),
            [],
        ).append(occurrence)

    for (provider_kind, provider_id, conflict_date, start_time), items in provider_slots.items():
        if len(items) <= 1:
            continue
        conflict_kind = (
            ConflictKind.STUDENT_DOUBLE_BOOKING
            if provider_kind == ReplacementProviderKind.STUDENT
            else ConflictKind.THERAPIST_DOUBLE_BOOKING
        )
        conflicts.append(
            ScheduleConflict(
                kind=conflict_kind,
                subject_id=provider_id,
                conflict_date=conflict_date,
                start_time=start_time,
                occurrence_ids=tuple(sorted(item.source_id for item in items)),
            )
        )

    for (patient_id, conflict_date, start_time), items in patient_slots.items():
        if len(items) > 1:
            conflicts.append(
                ScheduleConflict(
                    kind=ConflictKind.PATIENT_DOUBLE_BOOKING,
                    subject_id=patient_id,
                    conflict_date=conflict_date,
                    start_time=start_time,
                    occurrence_ids=tuple(sorted(item.source_id for item in items)),
                )
            )

    conflicts.sort(
        key=lambda conflict: (
            conflict.conflict_date,
            conflict.start_time,
            conflict.kind.value,
            conflict.subject_id,
        )
    )
    return conflicts
