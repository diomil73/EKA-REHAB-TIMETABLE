from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from .day_patterns import patterns_overlap
from .models import BaseScheduleEntry, Patient


@dataclass(frozen=True)
class RecurringTherapistConflict:
    therapist_id: str
    start_time: time
    existing_patient_id: str
    existing_patient_name: str
    new_patient_id: str
    new_patient_name: str
    existing_day_pattern: str
    new_day_pattern: str
    existing_base_entry_id: str


def find_recurring_therapist_conflicts(
    *,
    patient_id: str,
    therapist_id: str | None,
    start_time: time,
    day_pattern: str,
    existing_entries: Iterable[BaseScheduleEntry],
    patients: Iterable[Patient] = (),
    ignore_base_entry_id: str | None = None,
) -> tuple[RecurringTherapistConflict, ...]:
    """Find recurring double bookings for one therapist/time/day-pattern request."""

    provider = (therapist_id or "").strip()
    if not provider:
        return ()

    patient_key = patient_id.strip()
    patient_names = {
        patient.patient_id: patient.display_name
        for patient in patients
    }
    new_patient_name = patient_names.get(patient_key, patient_key)

    conflicts: list[RecurringTherapistConflict] = []
    for entry in existing_entries:
        if ignore_base_entry_id and entry.base_entry_id == ignore_base_entry_id:
            continue
        if not entry.therapist_id:
            continue
        if entry.therapist_id.strip().casefold() != provider.casefold():
            continue
        if entry.start_time != start_time:
            continue
        if not patterns_overlap(day_pattern, entry.day_pattern):
            continue

        conflicts.append(
            RecurringTherapistConflict(
                therapist_id=provider,
                start_time=start_time,
                existing_patient_id=entry.patient_id,
                existing_patient_name=patient_names.get(entry.patient_id, entry.patient_id),
                new_patient_id=patient_key,
                new_patient_name=new_patient_name,
                existing_day_pattern=entry.day_pattern,
                new_day_pattern=day_pattern,
                existing_base_entry_id=entry.base_entry_id,
            )
        )

    return tuple(
        sorted(
            conflicts,
            key=lambda item: (
                item.start_time,
                item.existing_patient_name.casefold(),
                item.existing_patient_id,
                item.existing_base_entry_id,
            ),
        )
    )
