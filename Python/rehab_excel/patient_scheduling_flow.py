from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from pathlib import Path
from typing import Iterable

from rehab_core.scheduling_candidates import TherapistScheduleCandidate

from .patient_treatment_plan import (
    PatientTreatmentPlan,
    TreatmentSchedulingOptions,
    build_patient_treatment_plan,
)


@dataclass(frozen=True)
class TherapistChoice:
    therapist_id: str
    label: str
    available_slots: tuple[time, ...]


@dataclass(frozen=True)
class TreatmentAssignment:
    treatment: str
    therapist_id: str
    start_time: time


@dataclass(frozen=True)
class AssignmentValidation:
    valid: bool
    message: str = ""
    refreshed_plan: PatientTreatmentPlan | None = None


def _normalize(value: str) -> str:
    return value.strip().casefold()


def _treatment_options(
    plan: PatientTreatmentPlan,
    treatment: str,
) -> TreatmentSchedulingOptions:
    key = _normalize(treatment)
    for item in plan.treatments:
        if _normalize(item.treatment) == key:
            return item
    raise ValueError(f"Treatment {treatment!r} is not part of the patient plan")


def therapist_choices(
    plan: PatientTreatmentPlan,
    treatment: str,
) -> tuple[TherapistChoice, ...]:
    """Return UI-ready therapist choices in the engine's workload order."""

    options = _treatment_options(plan, treatment)
    if not options.requires_therapist:
        return ()
    return tuple(
        TherapistChoice(
            therapist_id=candidate.therapist_id,
            label=candidate.workload_label,
            available_slots=candidate.available_slots,
        )
        for candidate in options.therapist_candidates
    )


def available_slots_for_therapist(
    plan: PatientTreatmentPlan,
    treatment: str,
    therapist_id: str,
) -> tuple[time, ...]:
    """Return only conflict-free slots for the selected therapist."""

    options = _treatment_options(plan, treatment)
    therapist_key = _normalize(therapist_id)
    for candidate in options.therapist_candidates:
        if _normalize(candidate.therapist_id) == therapist_key:
            return candidate.available_slots
    return ()


def validate_assignment_against_plan(
    plan: PatientTreatmentPlan,
    assignment: TreatmentAssignment,
) -> AssignmentValidation:
    """Validate one UI selection against the current plan snapshot."""

    options = _treatment_options(plan, assignment.treatment)
    if not options.requires_therapist:
        return AssignmentValidation(
            valid=False,
            message=f"Treatment {options.treatment!r} is not therapist-scheduled by this flow",
        )

    therapist_key = _normalize(assignment.therapist_id)
    for candidate in options.therapist_candidates:
        if _normalize(candidate.therapist_id) != therapist_key:
            continue
        if assignment.start_time in candidate.available_slots:
            return AssignmentValidation(valid=True)
        return AssignmentValidation(
            valid=False,
            message="The selected time is no longer available for this therapist",
        )

    return AssignmentValidation(
        valid=False,
        message="The selected therapist is no longer available for this treatment",
    )


def refresh_and_validate_assignment(
    workbook_path: str | Path,
    *,
    patient_id: str,
    selected_treatments: Iterable[str],
    target_date: date,
    assignment: TreatmentAssignment,
    additional_robotic_capable_names: Iterable[str] = (),
) -> AssignmentValidation:
    """Re-read workbook state and validate immediately before commit.

    This is the realtime safety boundary for the UI: therapist and time choices
    are treated as provisional until the authoritative workbook is re-read.
    """

    refreshed = build_patient_treatment_plan(
        workbook_path,
        patient_id=patient_id,
        selected_treatments=selected_treatments,
        target_date=target_date,
        additional_robotic_capable_names=additional_robotic_capable_names,
    )
    result = validate_assignment_against_plan(refreshed, assignment)
    return AssignmentValidation(
        valid=result.valid,
        message=result.message,
        refreshed_plan=refreshed,
    )
