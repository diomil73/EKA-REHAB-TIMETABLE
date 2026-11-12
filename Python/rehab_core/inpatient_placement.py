"""Effective-dated inpatient placement and care-team changes.

A PatientID is a stable clinical identity. Room, responsible doctor, and
treatment assignments are mutable *attributes*, not identity keys.

This pure module records no patient names, never changes history or a workbook,
and deliberately separates ward/doctor transitions from therapy scheduling.
Native Excel persistence and immutable historical snapshot integration remain
separate guarded workflows.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date
from typing import Iterable

from .models import Patient, PatientType


class InpatientPlacementError(ValueError):
    """Unsafe effective-dated room/doctor transition."""


@dataclass(frozen=True)
class InpatientPlacementChange:
    patient_id: str
    effective_date: date
    new_room: str | None = None
    new_responsible_doctor: str | None = None
    reason: str = ""

    def __post_init__(self) -> None:
        if not self.patient_id.strip():
            raise InpatientPlacementError("Stable PatientID is required")
        if self.new_room is None and self.new_responsible_doctor is None:
            raise InpatientPlacementError("At least one placement attribute must change")
        if self.new_room is not None and not self.new_room.strip():
            raise InpatientPlacementError("New room cannot be empty")
        if self.new_responsible_doctor is not None and not self.new_responsible_doctor.strip():
            raise InpatientPlacementError("New responsible doctor cannot be empty")


@dataclass(frozen=True)
class InpatientPlacementState:
    patient_id: str
    effective_on: date
    room: str | None
    responsible_doctor: str | None


def placement_on_date(
    patient: Patient,
    *,
    baseline_date: date,
    target_date: date,
    changes: Iterable[InpatientPlacementChange] = (),
) -> InpatientPlacementState:
    """Resolve a ward/doctor state without changing patient or scheduled therapy.

    Events are sorted by effective date. Only one event per patient/date is
    accepted, to avoid ambiguous same-day corrections. An update to one field
    leaves the other field untouched. Old-day queries retain old-day state.
    """
    if patient.patient_type != PatientType.INPATIENT:
        raise InpatientPlacementError("Placement transitions require an inpatient")
    if not patient.patient_id.strip():
        raise InpatientPlacementError("Stable PatientID is required")
    if target_date < baseline_date:
        raise InpatientPlacementError("Cannot infer placement before known baseline")
    events = tuple(changes)
    if any(event.patient_id != patient.patient_id for event in events):
        raise InpatientPlacementError("Placement event references a different PatientID")
    if any(event.effective_date < baseline_date for event in events):
        raise InpatientPlacementError("Placement event predates known baseline")
    days = [event.effective_date for event in events]
    if len(days) != len(set(days)):
        raise InpatientPlacementError("Multiple placement events for the same day")
    room, doctor = patient.room, patient.responsible_doctor
    for event in sorted(events, key=lambda item: item.effective_date):
        if event.effective_date > target_date:
            break
        if event.new_room is not None:
            room = event.new_room
        if event.new_responsible_doctor is not None:
            doctor = event.new_responsible_doctor
    return InpatientPlacementState(patient.patient_id, target_date, room, doctor)


def patient_with_current_placement(
    patient: Patient,
    *,
    baseline_date: date,
    target_date: date,
    changes: Iterable[InpatientPlacementChange] = (),
) -> Patient:
    """Produce a current roster projection, keeping patient identity unchanged."""
    state = placement_on_date(
        patient, baseline_date=baseline_date,
        target_date=target_date, changes=changes,
    )
    return replace(
        patient,
        room=state.room,
        responsible_doctor=state.responsible_doctor,
    )
