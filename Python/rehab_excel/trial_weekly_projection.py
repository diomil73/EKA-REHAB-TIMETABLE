"""Pure translation of an approved private trial plan into weekly sheet cells.

No workbook creation, COM, file I/O or patient-registry changes. The caller must
supply independently confirmed inpatient IDs; they are NEVER inferred.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import time
from typing import Iterable

from rehab_core.programme_mapping import SourceWeeklyEntry, validate_weekly_entry_consistency, ProgrammeMappingError
from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.weekly_slots_source import (
    HEADERS, KNOWN_TREATMENTS, REQUIRES_PROVIDER, read_weekly_slots_sheet,
    WeeklySlotsSourceError,
)


class TrialWeeklyProjectionError(ValueError):
    """A proposed weekly projection does not match approved source semantics."""


@dataclass(frozen=True)
class TrialWeeklyProjection:
    headers: tuple[str, ...]
    rows: tuple[tuple[str, str, str, time, str | None, str], ...] = field(repr=False)
    row_count: int = 0


def project_trial_weekly_rows(
    plan: TrialImportPlan, *, confirmed_inpatient_ids: Iterable[str]
) -> TrialWeeklyProjection:
    """Prepare exact PROGRAM_WEEKLY_SLOTS rows, without writing any Excel file.

    This only covers the weekly schedule sheet, not patient demographics or
    treatment admission into PATIENT_PLANNER. Never use as import authorization.
    """
    allowed_list = tuple(confirmed_inpatient_ids)
    allowed = set(allowed_list)
    if (
        not allowed_list or len(allowed) != len(allowed_list)
        or any(not isinstance(pid, str) or not pid.strip() for pid in allowed_list)
    ):
        raise TrialWeeklyProjectionError("Explicit unique inpatient IDs required")
    if not plan.entries or len(plan.entries) != plan.weekday_slot_count:
        raise TrialWeeklyProjectionError("Source plan slot count mismatch")

    try:
        validate_weekly_entry_consistency(list(plan.entries))
    except ProgrammeMappingError as exc:
        raise TrialWeeklyProjectionError("Source entries collide") from exc

    rows = []
    for entry in plan.entries:
        if not isinstance(entry, SourceWeeklyEntry):
            raise TrialWeeklyProjectionError("Unsupported source entry")
        if entry.patient_id not in allowed:
            raise TrialWeeklyProjectionError("Unconfirmed or non-inpatient ID")
        if entry.treatment not in KNOWN_TREATMENTS:
            raise TrialWeeklyProjectionError("Unknown treatment")
        if entry.treatment in REQUIRES_PROVIDER and not entry.therapist_name:
            raise TrialWeeklyProjectionError("Mandatory provider missing")
        rows.append((
            entry.patient_id, entry.treatment, entry.weekday,
            entry.start_time, entry.therapist_name, entry.source_cell,
        ))
    if len(rows) != plan.weekday_slot_count:
        raise TrialWeeklyProjectionError("Unexpected row count")
    return TrialWeeklyProjection(HEADERS, tuple(rows), len(rows))
