from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Iterable

from rehab_core.day_patterns import parse_day_pattern
from rehab_core.models import BaseScheduleEntry, Session, Therapist
from rehab_core.scheduling_candidates import (
    TherapistScheduleCandidate,
    rank_therapist_candidates,
)

from .reader import read_base_schedule, read_patients, read_settings
from .student_registry import read_students


THERAPIST_SCHEDULING_TREATMENTS = frozenset({"ΦΘ", "Ρομποτικό"})


def _normalize(value: str) -> str:
    return value.strip().casefold()


def _entry_applies_on(entry: BaseScheduleEntry, target_date: date) -> bool:
    return target_date.weekday() in {int(day) for day in parse_day_pattern(entry.day_pattern)}


def materialize_recurring_sessions(
    entries: Iterable[BaseScheduleEntry],
    target_date: date,
) -> tuple[Session, ...]:
    """Materialize provider-owned recurring entries for one concrete workday."""

    sessions: list[Session] = []
    for entry in entries:
        if not entry.therapist_id:
            continue
        if not _entry_applies_on(entry, target_date):
            continue
        sessions.append(
            Session(
                session_id=entry.base_entry_id,
                patient_id=entry.patient_id,
                therapist_id=entry.therapist_id,
                session_date=target_date,
                start_time=entry.start_time,
                treatment=entry.treatment,
                robotic=entry.robotic,
            )
        )
    return tuple(sessions)


def _robotic_capable_names(
    entries: Iterable[BaseScheduleEntry],
    additional_names: Iterable[str] = (),
) -> frozenset[str]:
    names = {
        _normalize(entry.therapist_id)
        for entry in entries
        if entry.robotic and entry.therapist_id
    }
    names.update(_normalize(name) for name in additional_names if name.strip())
    return frozenset(names)


def _workbook_therapists(
    therapist_names: Iterable[str],
    *,
    robotic_capable_names: Iterable[str] = (),
    excluded_names: Iterable[str] = (),
) -> tuple[Therapist, ...]:
    robotic = {_normalize(name) for name in robotic_capable_names}
    excluded = {_normalize(name) for name in excluded_names if name.strip()}
    result: list[Therapist] = []
    seen: set[str] = set()
    for raw_name in therapist_names:
        name = raw_name.strip()
        key = _normalize(name)
        if not name or key in seen or key in excluded:
            continue
        seen.add(key)
        # Until the workbook exposes a separate stable therapist id in SETTINGS,
        # the authoritative provider label is used as both id and display name.
        result.append(
            Therapist(
                therapist_id=name,
                display_name=name,
                robotic_capable=key in robotic,
            )
        )
    return tuple(result)


def workbook_therapist_candidates(
    workbook_path: str | Path,
    *,
    patient_id: str,
    treatment: str,
    target_date: date,
    additional_robotic_capable_names: Iterable[str] = (),
    ignore_session_id: str | None = None,
) -> tuple[TherapistScheduleCandidate, ...]:
    """Rank real workbook therapists and slots for ΦΘ/Ρομποτικό scheduling.

    SETTINGS provides the therapist order and standard slots, PATIENT_PLANNER
    supplies the recurring programme, PATIENTS validates the requested patient,
    and STUDENTS is used to prevent student display names from leaking into the
    physiotherapist candidate pool when legacy SETTINGS data contains them.

    Existing robotic assignments are treated as evidence of robotic capability;
    callers may provide additional confirmed names. Other treatment families use
    different provider/resource rules and are not silently routed through the
    physiotherapist engine.
    """

    treatment_text = treatment.strip()
    if treatment_text not in THERAPIST_SCHEDULING_TREATMENTS:
        raise ValueError(
            f"Treatment {treatment!r} is not handled by the physiotherapist scheduling engine"
        )

    patient_key = patient_id.strip()
    if not patient_key:
        raise ValueError("patient_id is required")

    settings = read_settings(workbook_path)
    entries = tuple(read_base_schedule(workbook_path))
    patients = tuple(read_patients(workbook_path))
    patient_ids = {_normalize(patient.patient_id) for patient in patients}
    if _normalize(patient_key) not in patient_ids:
        raise ValueError(f"Unknown PatientID: {patient_id}")

    students = tuple(read_students(workbook_path))
    student_names = tuple(student.display_name for student in students)
    sessions = materialize_recurring_sessions(entries, target_date)

    robotic_names = _robotic_capable_names(
        entries,
        additional_names=additional_robotic_capable_names,
    )
    therapists = _workbook_therapists(
        settings.therapist_names,
        robotic_capable_names=robotic_names,
        excluded_names=student_names,
    )

    return rank_therapist_candidates(
        patient_id=patient_key,
        treatment=treatment_text,
        target_date=target_date,
        timeslots=settings.standard_timeslots,
        therapists=therapists,
        sessions=sessions,
        patients=patients,
        ignore_session_id=ignore_session_id,
    )
