from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from pathlib import Path
from typing import Iterable

from .reader import read_patients, read_settings


PROVIDER_NONE = "none"
PROVIDER_PHYSIOTHERAPIST = "physiotherapist"
PROVIDER_OCCUPATIONAL_THERAPIST = "occupational_therapist"
PROVIDER_SPEECH_THERAPIST = "speech_therapist"
PROVIDER_PSYCHOLOGIST = "psychologist"


def _normalize(value: str) -> str:
    return value.strip().casefold()


_PROVIDER_BY_TREATMENT = {
    _normalize("ΦΘ"): PROVIDER_PHYSIOTHERAPIST,
    _normalize("Ρομποτικό"): PROVIDER_PHYSIOTHERAPIST,
    _normalize("Εργο"): PROVIDER_OCCUPATIONAL_THERAPIST,
    _normalize("Λογο"): PROVIDER_SPEECH_THERAPIST,
    _normalize("Ψυχ"): PROVIDER_PSYCHOLOGIST,
}


@dataclass(frozen=True)
class TreatmentProgramSpec:
    """One treatment exposed by SETTINGS in the patient-program editor."""

    treatment: str
    provider_kind: str
    provider_choices: tuple[str, ...] = ()

    @property
    def requires_provider(self) -> bool:
        return self.provider_kind != PROVIDER_NONE


@dataclass(frozen=True)
class TreatmentDraftConstraint:
    """User-supplied pieces of one selected treatment.

    Every non-None field is a fixed/manual constraint. Missing fields are free
    for automatic completion. This deliberately supports mixed workflows such
    as fixed days with automatic therapist/time, or a completely fixed row.
    """

    treatment: str
    day_pattern: str | None = None
    start_time: time | None = None
    provider_id: str | None = None

    @property
    def is_fully_manual(self) -> bool:
        return (
            self.day_pattern is not None
            and self.start_time is not None
            and self.provider_id is not None
        )


@dataclass(frozen=True)
class PatientProgramDraft:
    patient_id: str
    treatments: tuple[TreatmentProgramSpec, ...]
    constraints: tuple[TreatmentDraftConstraint, ...] = ()

    @property
    def treatment_names(self) -> tuple[str, ...]:
        return tuple(item.treatment for item in self.treatments)

    def spec_for(self, treatment: str) -> TreatmentProgramSpec:
        key = _normalize(treatment)
        for item in self.treatments:
            if _normalize(item.treatment) == key:
                return item
        raise KeyError(treatment)


def _provider_kind(treatment: str) -> str:
    return _PROVIDER_BY_TREATMENT.get(_normalize(treatment), PROVIDER_NONE)


def _provider_choices(settings, provider_kind: str) -> tuple[str, ...]:
    if provider_kind == PROVIDER_PHYSIOTHERAPIST:
        return tuple(settings.therapist_names)
    if provider_kind == PROVIDER_OCCUPATIONAL_THERAPIST:
        return tuple(settings.occupational_therapist_names)
    if provider_kind == PROVIDER_SPEECH_THERAPIST:
        return tuple(settings.speech_therapist_names)
    if provider_kind == PROVIDER_PSYCHOLOGIST:
        return tuple(settings.psychologist_names)
    return ()


def build_patient_program_draft(
    workbook_path: str | Path,
    *,
    patient_id: str,
    constraints: Iterable[TreatmentDraftConstraint] = (),
) -> PatientProgramDraft:
    """Build a read-only patient-program draft from authoritative workbook data.

    Every treatment in SETTINGS is exposed, in SETTINGS order. Known treatment
    families receive their appropriate provider pool; treatments without a
    provider remain fully selectable. The function never writes to the workbook.
    """

    patient_key = patient_id.strip()
    if not patient_key:
        raise ValueError("patient_id is required")

    patients = tuple(read_patients(workbook_path))
    if _normalize(patient_key) not in {_normalize(item.patient_id) for item in patients}:
        raise ValueError(f"Unknown PatientID: {patient_id}")

    settings = read_settings(workbook_path)
    specs: list[TreatmentProgramSpec] = []
    authoritative: dict[str, str] = {}
    for raw in settings.therapies:
        treatment = raw.strip()
        if not treatment:
            continue
        key = _normalize(treatment)
        if key in authoritative:
            continue
        authoritative[key] = treatment
        kind = _provider_kind(treatment)
        specs.append(
            TreatmentProgramSpec(
                treatment=treatment,
                provider_kind=kind,
                provider_choices=_provider_choices(settings, kind),
            )
        )

    day_patterns = {_normalize(item): item for item in settings.day_patterns}
    normalized_constraints: list[TreatmentDraftConstraint] = []
    seen: set[str] = set()
    for constraint in constraints:
        key = _normalize(constraint.treatment)
        if key not in authoritative:
            raise ValueError(f"Treatment {constraint.treatment!r} is not defined in SETTINGS")
        if key in seen:
            raise ValueError(f"Duplicate treatment constraint: {constraint.treatment}")
        seen.add(key)

        canonical_day = None
        if constraint.day_pattern is not None:
            day_key = _normalize(constraint.day_pattern)
            if day_key not in day_patterns:
                raise ValueError(
                    f"Day pattern {constraint.day_pattern!r} is not defined in SETTINGS"
                )
            canonical_day = day_patterns[day_key]

        spec = next(item for item in specs if _normalize(item.treatment) == key)
        provider = constraint.provider_id.strip() if constraint.provider_id else None
        if provider is not None:
            if not spec.requires_provider:
                raise ValueError(f"Treatment {spec.treatment!r} does not use a provider")
            allowed = {_normalize(item) for item in spec.provider_choices}
            if _normalize(provider) not in allowed:
                raise ValueError(
                    f"Provider {provider!r} is not defined for treatment {spec.treatment!r}"
                )
            provider = next(
                item for item in spec.provider_choices if _normalize(item) == _normalize(provider)
            )

        normalized_constraints.append(
            TreatmentDraftConstraint(
                treatment=authoritative[key],
                day_pattern=canonical_day,
                start_time=constraint.start_time,
                provider_id=provider,
            )
        )

    return PatientProgramDraft(
        patient_id=patient_key,
        treatments=tuple(specs),
        constraints=tuple(normalized_constraints),
    )
