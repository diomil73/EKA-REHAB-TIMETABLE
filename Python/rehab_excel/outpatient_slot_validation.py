from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from rehab_core.day_patterns import patterns_overlap
from rehab_core.models import BaseScheduleEntry, Patient


class OutpatientSlotConflictError(ValueError):
    """Raised when one Excel cell would need incompatible patient-type fills."""


@dataclass(frozen=True)
class MixedPatientTypeSlotConflict:
    therapist_id: str
    start_time: time
    outpatient_patient_id: str
    inpatient_patient_id: str
    outpatient_day_pattern: str
    inpatient_day_pattern: str
    inpatient_base_entry_id: str


def find_outpatient_inpatient_slot_conflicts(
    *,
    outpatient_patient_id: str,
    therapist_id: str | None,
    start_time: time,
    day_pattern: str,
    existing_entries: Iterable[BaseScheduleEntry],
    patients: Iterable[Patient],
) -> tuple[MixedPatientTypeSlotConflict, ...]:
    """Find same-cell weekday overlaps between one outpatient and inpatients.

    THERAPIST_DAILY uses one Excel cell per provider/time and therefore one
    background fill. Outpatient blue can coexist with an inpatient assignment
    in the same recurring cell only when their weekday patterns are
    complementary. Overlapping weekdays would require two incompatible fills
    in one cell and are rejected before write-back.
    """

    provider = (therapist_id or "").strip()
    if not provider:
        return ()

    patient_by_id = {patient.patient_id: patient for patient in patients}
    conflicts: list[MixedPatientTypeSlotConflict] = []

    for entry in existing_entries:
        if entry.therapist_id is None:
            continue
        if entry.therapist_id.strip().casefold() != provider.casefold():
            continue
        if entry.start_time != start_time:
            continue

        patient = patient_by_id.get(entry.patient_id)
        if patient is None or patient.is_outpatient:
            continue
        if not patterns_overlap(day_pattern, entry.day_pattern):
            continue

        conflicts.append(
            MixedPatientTypeSlotConflict(
                therapist_id=provider,
                start_time=start_time,
                outpatient_patient_id=outpatient_patient_id,
                inpatient_patient_id=entry.patient_id,
                outpatient_day_pattern=day_pattern,
                inpatient_day_pattern=entry.day_pattern,
                inpatient_base_entry_id=entry.base_entry_id,
            )
        )

    return tuple(conflicts)


def validate_outpatient_slot_compatibility(
    *,
    outpatient_patient_id: str,
    therapist_id: str | None,
    start_time: time,
    day_pattern: str,
    existing_entries: Iterable[BaseScheduleEntry],
    patients: Iterable[Patient],
) -> None:
    conflicts = find_outpatient_inpatient_slot_conflicts(
        outpatient_patient_id=outpatient_patient_id,
        therapist_id=therapist_id,
        start_time=start_time,
        day_pattern=day_pattern,
        existing_entries=existing_entries,
        patients=patients,
    )
    if not conflicts:
        return

    first = conflicts[0]
    raise OutpatientSlotConflictError(
        "Outpatient recurring slot overlaps an inpatient in the same "
        f"THERAPIST_DAILY cell: therapist {first.therapist_id!r}, "
        f"time {first.start_time.strftime('%H:%M')}, "
        f"outpatient days {first.outpatient_day_pattern!r}, "
        f"inpatient {first.inpatient_patient_id!r} days "
        f"{first.inpatient_day_pattern!r}. Use complementary weekdays or a "
        "different therapist/time slot."
    )
