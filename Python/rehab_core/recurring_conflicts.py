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


def _provider_key(value: str | None) -> str:
    text = (value or "").strip().casefold()
    replacements = {
        "ά": "α",
        "έ": "ε",
        "ή": "η",
        "ί": "ι",
        "ϊ": "ι",
        "ΐ": "ι",
        "ό": "ο",
        "ύ": "υ",
        "ϋ": "υ",
        "ΰ": "υ",
        "ώ": "ω",
        "ς": "σ",
        ".": "",
        " ": "",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


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
    """Find recurring double bookings for one therapist/time/day-pattern request.

    Workbook therapist labels are normalized with the same accent/spacing rules
    used by the native THERAPIST_DAILY mapping. This prevents a visible slot from
    being treated as a different provider merely because one source contains an
    accent, period or spacing variation.
    """

    provider = (therapist_id or "").strip()
    provider_key = _provider_key(provider)
    if not provider_key:
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
        if _provider_key(entry.therapist_id) != provider_key:
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
