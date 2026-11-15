from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable

from rehab_core.scheduling_candidates import TherapistScheduleCandidate

from .reader import read_settings
from .workbook_scheduling_candidates import (
    THERAPIST_SCHEDULING_TREATMENTS,
    workbook_therapist_candidates,
)


@dataclass(frozen=True)
class TreatmentSchedulingOptions:
    """Scheduling options for one selected patient treatment."""

    treatment: str
    requires_therapist: bool
    therapist_candidates: tuple[TherapistScheduleCandidate, ...] = ()


@dataclass(frozen=True)
class PatientTreatmentPlan:
    """Validated selected treatments and their current scheduling options."""

    patient_id: str
    target_date: date
    treatments: tuple[TreatmentSchedulingOptions, ...]

    @property
    def selected_treatments(self) -> tuple[str, ...]:
        return tuple(item.treatment for item in self.treatments)


def _normalized(value: str) -> str:
    return value.strip().casefold()


def build_patient_treatment_plan(
    workbook_path: str | Path,
    *,
    patient_id: str,
    selected_treatments: Iterable[str],
    target_date: date,
    additional_robotic_capable_names: Iterable[str] = (),
) -> PatientTreatmentPlan:
    """Validate treatment choices and expose current therapist/time candidates.

    The workbook SETTINGS sheet remains authoritative for which treatment labels
    are selectable. Treatments handled by the common physiotherapist scheduling
    engine (currently ΦΘ and Ρομποτικό) are immediately resolved into ranked,
    conflict-free therapist candidates and their available slots. Other treatment
    families remain in the same patient plan but do not get silently routed
    through the physiotherapist engine because they have different provider or
    resource semantics.
    """

    patient_key = patient_id.strip()
    if not patient_key:
        raise ValueError("patient_id is required")

    requested: list[str] = []
    seen: set[str] = set()
    for raw in selected_treatments:
        treatment = raw.strip()
        if not treatment:
            continue
        key = _normalized(treatment)
        if key in seen:
            raise ValueError(f"Duplicate treatment selection: {treatment}")
        seen.add(key)
        requested.append(treatment)

    if not requested:
        raise ValueError("At least one treatment must be selected")

    settings = read_settings(workbook_path)
    authoritative = {_normalized(item): item for item in settings.therapies}

    canonical: list[str] = []
    for treatment in requested:
        key = _normalized(treatment)
        if key not in authoritative:
            raise ValueError(f"Treatment {treatment!r} is not defined in SETTINGS")
        canonical.append(authoritative[key])

    therapist_treatments = {_normalized(item) for item in THERAPIST_SCHEDULING_TREATMENTS}
    options: list[TreatmentSchedulingOptions] = []

    for treatment in canonical:
        requires_therapist = _normalized(treatment) in therapist_treatments
        candidates: tuple[TherapistScheduleCandidate, ...] = ()
        if requires_therapist:
            candidates = workbook_therapist_candidates(
                workbook_path,
                patient_id=patient_key,
                treatment=treatment,
                target_date=target_date,
                additional_robotic_capable_names=additional_robotic_capable_names,
            )
        options.append(
            TreatmentSchedulingOptions(
                treatment=treatment,
                requires_therapist=requires_therapist,
                therapist_candidates=candidates,
            )
        )

    return PatientTreatmentPlan(
        patient_id=patient_key,
        target_date=target_date,
        treatments=tuple(options),
    )
