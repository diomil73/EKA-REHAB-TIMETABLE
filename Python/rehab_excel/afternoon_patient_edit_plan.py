"""Pure proposed afternoon-physiotherapy patient edit; NO workbook writes.

Edits only one explicitly selected patient's independent afternoon namespace.
An empty replacement is an explicit removal only with allow_clear=True.
All other patients' rows and all morning sheets remain outside this function.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from rehab_excel.afternoon_slots_source import (
    AfternoonSlot,
    AfternoonSlotsError,
    validate_afternoon_rows,
)


class AfternoonEditError(ValueError):
    """An afternoon schedule proposal cannot be applied safely."""


@dataclass(frozen=True)
class AfternoonEditPlan:
    """Immutable, read-only proposed replacement of afternoon rows."""

    patient_id: str
    before: tuple[AfternoonSlot, ...]
    after: tuple[AfternoonSlot, ...]
    removed_count: int
    added_count: int
    unchanged_count: int

    @property
    def changes_required(self) -> bool:
        return self.before != self.after


def prepare_afternoon_patient_edit(
    current_rows: Iterable[AfternoonSlot],
    *,
    patient_id: str,
    replacement_rows: Iterable[AfternoonSlot],
    known_patient_ids: Iterable[str],
    allow_clear: bool = False,
) -> AfternoonEditPlan:
    """Produce a full-sheet candidate, never a side-effect or save permission.

    Existing order for other patients is unchanged. Replacement rows are placed
    where the patient's old rows began, or appended for a new patient. Repeated
    input is stable. Do not guess consent to delete when the replacement is
    empty. A later native writer needs source hashing, protection handling,
    atomic read-back and a separate confirmation step.
    """
    if not isinstance(patient_id, str) or not patient_id.strip() or patient_id != patient_id.strip():
        raise AfternoonEditError("Canonical explicit patient ID required")
    confirmed = tuple(known_patient_ids)
    if not confirmed or len(set(confirmed)) != len(confirmed) or any(
        not isinstance(x, str) or not x.strip() for x in confirmed
    ) or patient_id not in confirmed:
        raise AfternoonEditError("Explicit unique confirmed patient IDs required")

    before = tuple(current_rows)
    replacements = tuple(replacement_rows)
    try:
        validate_afternoon_rows(before, known_patient_ids=confirmed)
        validate_afternoon_rows(replacements, known_patient_ids=confirmed)
    except AfternoonSlotsError as exc:
        raise AfternoonEditError("Invalid existing or proposed afternoon rows") from exc

    if any(item.patient_id != patient_id for item in replacements):
        raise AfternoonEditError("Replacement includes a different patient")
    old_patient_rows = tuple(item for item in before if item.patient_id == patient_id)
    if old_patient_rows and not replacements and not allow_clear:
        raise AfternoonEditError("Explicit patient clear confirmation required")

    insert_at = next(
        (index for index, item in enumerate(before) if item.patient_id == patient_id),
        len(before),
    )
    other_rows = tuple(item for item in before if item.patient_id != patient_id)
    safe_insert_at = sum(item.patient_id != patient_id for item in before[:insert_at])
    after = (
        other_rows[:safe_insert_at] + replacements + other_rows[safe_insert_at:]
    )

    try:
        validate_afternoon_rows(after, known_patient_ids=confirmed)
    except AfternoonSlotsError as exc:
        raise AfternoonEditError("Combined afternoon proposal is invalid") from exc

    original = set(old_patient_rows)
    proposed = set(replacements)
    return AfternoonEditPlan(
        patient_id=patient_id,
        before=before,
        after=after,
        removed_count=len(original - proposed),
        added_count=len(proposed - original),
        unchanged_count=len(original & proposed),
    )
