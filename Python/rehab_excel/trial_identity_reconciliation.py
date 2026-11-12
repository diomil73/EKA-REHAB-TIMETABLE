"""Read-only identity reconciliation between new PROGRAMMA and old CC patients.

Never auto-confirm an ID. A unique exact normalized name can be proposed only;
duplicates, name variants and collisions always require manual verification.
No clinical records are written and no real names appear in logs.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Iterable
import unicodedata


@dataclass(frozen=True)
class PatientIdentity:
    source_row: int
    clinic: str
    name: str


@dataclass(frozen=True)
class OldPatientIdentity:
    patient_id: str
    name: str


@dataclass(frozen=True)
class IdentityCandidate:
    source_row: int
    clinic: str
    suggested_patient_id: str | None
    state: str


def normalize_patient_name(value: str) -> str:
    text = unicodedata.normalize("NFD", value.strip().casefold())
    unaccented = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return " ".join(unaccented.split())


def suggest_identity_matches(
    current: Iterable[PatientIdentity],
    old: Iterable[OldPatientIdentity],
) -> tuple[IdentityCandidate, ...]:
    """Only one-to-one exact-name matches become *unconfirmed suggestions*.

    All old IDs are protected from recycling; an identifier in a candidate is
    informational until the operator explicitly verifies the actual person.
    """
    source = tuple(current)
    history = tuple(old)
    old_by_name: dict[str, list[str]] = {}
    source_names = Counter(normalize_patient_name(item.name) for item in source)
    source_locations = Counter((item.clinic.strip(), item.source_row) for item in source)
    for person in history:
        if not person.patient_id.strip() or not person.name.strip():
            continue
        old_by_name.setdefault(normalize_patient_name(person.name), []).append(
            person.patient_id.strip()
        )
    all_old_ids = Counter(person.patient_id.strip().casefold() for person in history)
    result = []
    for patient in source:
        name = normalize_patient_name(patient.name)
        candidates = old_by_name.get(name, ())
        if not name or not patient.clinic.strip() or patient.source_row < 2:
            state, suggestion = "INVALID_SOURCE_IDENTITY", None
        elif source_locations[(patient.clinic.strip(), patient.source_row)] != 1:
            state, suggestion = "DUPLICATE_SOURCE_LOCATION", None
        elif source_names[name] != 1:
            state, suggestion = "AMBIGUOUS_CURRENT_NAME", None
        elif len(candidates) != 1 or all_old_ids[candidates[0].casefold()] != 1:
            state, suggestion = (
                ("NO_EXACT_OLD_MATCH" if not candidates else "AMBIGUOUS_OLD_IDENTITY"),
                None,
            )
        else:
            state, suggestion = "EXACT_NAME_CANDIDATE_UNCONFIRMED", candidates[0]
        result.append(IdentityCandidate(
            source_row=patient.source_row,
            clinic=patient.clinic,
            suggested_patient_id=suggestion,
            state=state,
        ))
    return tuple(result)
