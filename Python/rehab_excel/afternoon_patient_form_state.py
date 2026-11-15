"""Read-only afternoon ΦΘ state for the existing patient-management flow.

The morning form uses PATIENT_PLANNER and its own per-treatment dictionary.
This adapter deliberately reads ONLY the independent afternoon technical sheet.
It never opens the workbook for editing or changes existing patient forms.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from rehab_excel.afternoon_slots_source import (
    AfternoonSlot,
    AfternoonSlotsError,
    SHEET_NAME,
    read_afternoon_slots_sheet,
)
from rehab_excel.patient_registry_source import read_patient_registry


class AfternoonFormStateError(ValueError):
    """Afternoon patient-management state is unavailable or invalid."""


@dataclass(frozen=True)
class AfternoonPatientFormState:
    """Private state for one selected patient; no identifiable repr."""

    patient_id: str
    enabled: bool
    slots: tuple[AfternoonSlot, ...]

    def __repr__(self) -> str:
        return (
            "AfternoonPatientFormState("
            f"enabled={self.enabled}, slot_count={len(self.slots)})"
        )


def read_afternoon_patient_form_state(
    workbook_path: str | Path, *, patient_id: str
) -> AfternoonPatientFormState:
    """Read after explicit patient choice without seeding anything.

    A legacy workbook without the PM sheet is reported as not installed, not
    as a usable blank writer. A newly installed header-only sheet yields an
    enabled but empty PM state. An invalid PM sheet always stops the operation.
    """
    if not isinstance(patient_id, str) or not patient_id.strip() or patient_id != patient_id.strip():
        raise AfternoonFormStateError("Canonical patient selection required")
    source = Path(workbook_path)
    if source.suffix.casefold() not in {".xlsx", ".xlsm"} or not source.is_file():
        raise AfternoonFormStateError("Existing Excel workbook required")
    try:
        registered = read_patient_registry(source)
        confirmed_ids = tuple(item.patient_id for item in registered)
        if len(set(confirmed_ids)) != len(confirmed_ids) or patient_id not in confirmed_ids:
            raise AfternoonFormStateError("Patient not present in confirmed registry")

        book = load_workbook(
            source, read_only=True, data_only=True,
            keep_vba=source.suffix.casefold() == ".xlsm",
        )
        try:
            if SHEET_NAME not in book.sheetnames:
                return AfternoonPatientFormState(patient_id, False, ())
            all_rows = read_afternoon_slots_sheet(
                book[SHEET_NAME], known_patient_ids=confirmed_ids
            )
            patient_rows = tuple(
                item for item in all_rows if item.patient_id == patient_id
            )
            return AfternoonPatientFormState(patient_id, True, patient_rows)
        finally:
            book.close()
    except (AfternoonFormStateError, AfternoonSlotsError):
        raise
    except Exception as exc:
        raise AfternoonFormStateError(
            "Afternoon state could not be verified; no changes made"
        ) from exc


def afternoon_form_status(state: AfternoonPatientFormState) -> str:
    """Greek user-facing status, without patient names or hidden writes."""
    if not isinstance(state, AfternoonPatientFormState):
        raise AfternoonFormStateError("Validated afternoon form state required")
    if not state.enabled:
        return "Απογευματινή ΦΘ: δεν έχει εγκατασταθεί ακόμη."
    if not state.slots:
        return "Απογευματινή ΦΘ: κενή. Δεν έχουν καταχωριστεί συνεδρίες."
    return f"Απογευματινή ΦΘ: {len(state.slots)} καταχωρισμένες συνεδρίες."


def afternoon_form_rows(state: AfternoonPatientFormState) -> tuple[tuple[str, str, str], ...]:
    """Return only explicit time/day/therapist values for patient UI."""
    if not isinstance(state, AfternoonPatientFormState):
        raise AfternoonFormStateError("Validated afternoon form state required")
    return tuple(
        (item.weekday, item.start_time.strftime("%H:%M"), item.therapist)
        for item in state.slots
    )
