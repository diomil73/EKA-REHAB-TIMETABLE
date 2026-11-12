"""Read-only Excel-to-core bridge for effective-dated recurring therapy changes.

This is a *proposal* only. It consumes the current Excel schedule and explicit
capacity facts, then runs the shared core preflight. It NEVER changes the
workbook, admits a new patient, saves history or invokes VBA.

The existing row-derived "weekly:N"/"outpatient:N" keys are NOT durable across
row insertions. Every proposal must therefore carry an exact previously
observed BaseScheduleEntry fingerprint and fail closed if Excel has changed.
A real committed event log will require immutable entry IDs in a later schema.
"""
from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from typing import Mapping, Iterable

from openpyxl import load_workbook

from rehab_core.models import BaseScheduleEntry, Therapist
from rehab_core.recurring_therapy_changes import (
    RecurringTherapyChange,
    RecurringTherapyChangeError,
    ValidatedRecurringChangePreview,
    prepare_validated_recurring_change,
)
from .outpatient_schedule_source import read_unified_base_schedule
from .patient_registry_source import read_patient_registry
from .reader import read_settings
from .stable_recurring_ids import bind_stable_assignment_ids, is_stable_assignment_id
from .weekly_slots_source import SHEET_NAME, overlay_weekly_entries, read_weekly_slots_sheet


class RecurringExcelPreviewError(ValueError):
    """Workbook schedule or assignment changed; proposal must be reconsidered."""


@dataclass(frozen=True)
class RecurringExcelPreview:
    """Ephemeral, non-persisted read-only preview bound to current workbook."""

    workbook_sha256: str
    base_entry_count: int
    proposed: ValidatedRecurringChangePreview


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def read_effective_base_entries(path: Path) -> tuple[BaseScheduleEntry, ...]:
    """Apply the optional weekly-granular overlay before checking a proposal."""
    patients = read_patient_registry(path)
    entries = read_unified_base_schedule(path, patients=patients)
    book = load_workbook(
        path, read_only=True, data_only=True,
        keep_vba=path.suffix.lower() == ".xlsm",
    )
    try:
        if SHEET_NAME in book.sheetnames:
            weekly = read_weekly_slots_sheet(
                book[SHEET_NAME],
                inpatient_ids=(p.patient_id for p in patients if not p.is_outpatient),
            )
            entries = overlay_weekly_entries(
                planner_entries=entries, weekly_entries=weekly,
            )
    finally:
        book.close()
    return tuple(entries)


def preview_recurring_change_from_excel(
    path: Path,
    *,
    baseline_date,
    change: RecurringTherapyChange,
    expected_entry: BaseScheduleEntry,
    provider_max_daily_slots: Mapping[str, int],
    current_changes: Iterable[RecurringTherapyChange] = (),
    require_stable_identity: bool = False,
) -> RecurringExcelPreview:
    """Read, verify stale-edit fingerprint, preflight, and verify unchanged bytes.

    The caller MUST supply actually approved per-therapist daily capacities.
    Never infer 6 slots for everyone from the default Therapist dataclass.
    """
    workbook = Path(path).resolve()
    if not workbook.is_file() or workbook.suffix.lower() != ".xlsm":
        raise RecurringExcelPreviewError("Existing XLSM copy required")
    if change.base_entry_id != expected_entry.base_entry_id:
        raise RecurringExcelPreviewError("Change ID differs from displayed assignment")
    before = _digest(workbook)
    try:
        entries = read_effective_base_entries(workbook)
        if require_stable_identity:
            entries = bind_stable_assignment_ids(workbook, entries)
            if not is_stable_assignment_id(change.base_entry_id):
                raise RecurringExcelPreviewError(
                    "Persistent proposal requires immutable assignment ID"
                )
        matching = [entry for entry in entries
                    if entry.base_entry_id == expected_entry.base_entry_id]
        if len(matching) != 1 or matching[0] != expected_entry:
            raise RecurringExcelPreviewError(
                "Assignment changed or row identity shifted since operator preview"
            )
        settings = read_settings(workbook)
        allowed_names = set(settings.therapist_names)
        if not allowed_names:
            raise RecurringExcelPreviewError("No verified therapist registry")
        # Unknown capacity must fail closed; 4-slot therapists are not allowed
        # to inherit a generic 6-slot limit.
        if not provider_max_daily_slots or set(provider_max_daily_slots) != allowed_names:
            raise RecurringExcelPreviewError(
                "Approved daily capacity required for every registered therapist"
            )
        if any(type(cap) is not int or cap < 1
               for cap in provider_max_daily_slots.values()):
            raise RecurringExcelPreviewError("Invalid therapist capacity")
        therapists = tuple(
            Therapist(therapist_id=name, display_name=name,
                      max_daily_timeslots=provider_max_daily_slots[name])
            for name in settings.therapist_names
        )
        preview = prepare_validated_recurring_change(
            entries, baseline_date=baseline_date, change=change,
            current_changes=current_changes, therapists=therapists,
        )
        return RecurringExcelPreview(
            workbook_sha256=before,
            base_entry_count=len(entries),
            proposed=preview,
        )
    except (RecurringTherapyChangeError, ValueError):
        raise
    finally:
        if not workbook.is_file() or _digest(workbook) != before:
            raise RecurringExcelPreviewError(
                "Source XLSM changed during read-only recurring proposal"
            )
