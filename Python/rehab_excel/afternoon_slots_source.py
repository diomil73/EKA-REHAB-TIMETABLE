"""Isolated afternoon physiotherapy weekly-row contract.

This is an in-memory schema/validator, not an Excel installer or persistence
operation. Never merge these rows into morning PATIENT_PLANNER treatment keys.
No real identities appear in tests or diagnostic output.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from rehab_core.patient_programme_draft import AFTERNOON_TREATMENT, AfternoonPhysioDraft
from rehab_core.programme_mapping import WEEKDAYS

SHEET_NAME = "AFTERNOON_PHYSIO_SLOTS"
HEADERS = ("PatientID", "Treatment", "Weekday", "StartTime", "Therapist")
PROGRAMME = "AFTERNOON"


class AfternoonSlotsError(ValueError):
    """A separately scoped afternoon programme is not safe to accept."""


@dataclass(frozen=True)
class AfternoonSlot:
    patient_id: str
    treatment: str
    weekday: str
    start_time: time
    therapist: str

    def __post_init__(self) -> None:
        if (
            not isinstance(self.patient_id, str)
            or not self.patient_id.strip()
            or self.treatment != AFTERNOON_TREATMENT
            or self.weekday not in WEEKDAYS
            or not isinstance(self.start_time, time)
            or not isinstance(self.therapist, str)
            or not self.therapist.strip()
        ):
            raise AfternoonSlotsError(
                "Afternoon row requires patient ID, ΦΘ, weekday, time and physiotherapist"
            )


def validate_afternoon_rows(
    rows: Iterable[AfternoonSlot], *, known_patient_ids: Iterable[str] | None = None
) -> tuple[AfternoonSlot, ...]:
    """Fail closed on duplicate patient time/day; preserve shared-session review.

    A shared therapist time across *different* patients is NOT automatically
    rejected because group sessions require a separate explicit approval policy.
    This validation never imports, deduplicates or changes the morning programme.
    """
    items = tuple(rows)
    known = set(known_patient_ids) if known_patient_ids is not None else None
    if known is not None and (not known or any(not isinstance(x, str) or not x.strip() for x in known)):
        raise AfternoonSlotsError("Confirmed patient registry required")
    seen: set[tuple[str, str, time]] = set()
    for item in items:
        if not isinstance(item, AfternoonSlot):
            raise AfternoonSlotsError("Unexpected afternoon row")
        if known is not None and item.patient_id not in known:
            raise AfternoonSlotsError("Unknown patient ID in afternoon schedule")
        key = (item.patient_id, item.weekday, item.start_time)
        if key in seen:
            raise AfternoonSlotsError("Patient has two afternoon slots at identical day/time")
        seen.add(key)
    return items


def make_afternoon_rows(
    patient_id: str, drafts: Iterable[AfternoonPhysioDraft]
) -> tuple[AfternoonSlot, ...]:
    """Convert explicit completed drafts; blank drafts create no appointments."""
    if not isinstance(patient_id, str) or not patient_id.strip():
        raise AfternoonSlotsError("Explicit patient ID required")
    result = []
    for draft in drafts:
        if not isinstance(draft, AfternoonPhysioDraft):
            raise AfternoonSlotsError("Unexpected afternoon draft")
        if draft.blank:
            continue
        result.append(
            AfternoonSlot(
                patient_id.strip(), AFTERNOON_TREATMENT,
                draft.weekday, draft.start_time, draft.physiotherapist,
            )
        )
    return validate_afternoon_rows(result)


def read_afternoon_slots_sheet(
    worksheet, *, known_patient_ids: Iterable[str]
) -> tuple[AfternoonSlot, ...]:
    """Read-only optional Excel sheet adapter, without creating the sheet.

    Header-only is valid and means no afternoon sessions. Unknown rows halt
    the operation; empty cells do not imply a default time/day/provider.
    """
    header = tuple(worksheet.cell(1, col).value for col in range(1, 6))
    if header != HEADERS:
        raise AfternoonSlotsError("Afternoon sheet headers do not match contract")
    rows = []
    for row_number, values in enumerate(
        worksheet.iter_rows(min_row=2, max_col=5, values_only=True), start=2
    ):
        if all(value is None or str(value).strip() == "" for value in values):
            continue
        try:
            patient_id, treatment, weekday, clock, therapist = values
            rows.append(
                AfternoonSlot(
                    patient_id=patient_id,
                    treatment=treatment,
                    weekday=weekday,
                    start_time=clock,
                    therapist=therapist,
                )
            )
        except (AfternoonSlotsError, ValueError, TypeError) as exc:
            raise AfternoonSlotsError(
                f"Invalid afternoon schedule row {row_number}"
            ) from exc
    return validate_afternoon_rows(rows, known_patient_ids=known_patient_ids)
