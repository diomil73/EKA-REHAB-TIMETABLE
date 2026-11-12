"""Read-only, source-bound proposal to APPEND one permanent therapy event.

This module NEVER writes an Excel file. Native Excel append, authorized
capacity allocation, rollback and clinical approval are separate future gates.
A proposal cannot be committed unless its source SHA-256 still matches.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from hashlib import sha256
from pathlib import Path
from typing import Mapping

from openpyxl import load_workbook

from rehab_core.models import BaseScheduleEntry
from rehab_core.recurring_therapy_changes import RecurringTherapyChange
from .recurring_change_event_source import (
    HEADERS, SHEET_NAME, _text, read_recurring_change_events,
)
from .recurring_change_preview import (
    preview_recurring_change_from_excel, read_effective_base_entries,
)
from .stable_recurring_ids import (
    bind_stable_assignment_ids, is_stable_assignment_id,
)


class RecurringAppendPlanError(ValueError):
    """Append proposal is stale, incomplete, or not safely append-only."""


@dataclass(frozen=True)
class RecurringAppendPlan:
    """Never an authorization to write to this or another workbook."""
    source_sha256: str
    event_id: str
    row_values: tuple[object, ...]
    event_count_before: int
    base_entry_count: int

    def __post_init__(self):
        if len(self.row_values) != len(HEADERS):
            raise RecurringAppendPlanError("Event row width is invalid")


def file_digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def verify_unchanged_source(path: Path, plan: RecurringAppendPlan) -> None:
    """Required immediately BEFORE a future native copy or transaction."""
    if not path.is_file() or file_digest(path) != plan.source_sha256:
        raise RecurringAppendPlanError("Source changed since event proposal")


def _event_ids_and_shape(path: Path) -> tuple[set[str], bool]:
    """Inspect EventID uniqueness, schema and trailing rows without saving."""
    wb = load_workbook(
        path, read_only=True, data_only=True,
        keep_vba=path.suffix.lower() == ".xlsm",
    )
    try:
        if SHEET_NAME not in wb.sheetnames:
            return set(), False
        sheet = wb[SHEET_NAME]
        if tuple(sheet.cell(1, i).value for i in range(1, len(HEADERS)+1)) != HEADERS:
            raise RecurringAppendPlanError("Event schema differs")
        if sheet.max_column > len(HEADERS):
            raise RecurringAppendPlanError("Unexpected event columns")
        event_ids = set()
        for row in sheet.iter_rows(min_row=2, max_col=len(HEADERS), values_only=True):
            if all(value is None or str(value).strip() == "" for value in row):
                raise RecurringAppendPlanError("Gaps in append-only event ledger")
            event_id = _text(row[0])
            if not event_id or event_id in event_ids:
                raise RecurringAppendPlanError("Missing or duplicate existing event ID")
            event_ids.add(event_id)
        return event_ids, True
    finally:
        wb.close()


def plan_recurring_event_append(
    workbook_path: Path,
    *,
    event_id: str,
    change: RecurringTherapyChange,
    expected_entry: BaseScheduleEntry,
    baseline_date: date,
    provider_max_daily_slots: Mapping[str, int],
) -> RecurringAppendPlan:
    """Preflight strict source-bounded append with genuine current events.

    Capacity values must be explicit approved values, one per known therapist.
    No existing event may be silently discarded, moved or backdated.
    """
    path = Path(workbook_path).resolve()
    if not path.is_file() or path.suffix.lower() != ".xlsm":
        raise RecurringAppendPlanError("Existing isolated XLSM required")
    if not event_id.startswith("change:") or len(event_id) < 16:
        raise RecurringAppendPlanError("Invalid opaque event ID")
    if not is_stable_assignment_id(change.base_entry_id):
        raise RecurringAppendPlanError("Immutable AssignmentID required")
    before = file_digest(path)
    try:
        rows = read_effective_base_entries(path)
        stable = bind_stable_assignment_ids(path, rows)
        original = next((entry for entry in stable
                         if entry.base_entry_id == change.base_entry_id), None)
        if original is None or original != expected_entry:
            raise RecurringAppendPlanError("Stale assignment fingerprint")
        ids, _sheet_present = _event_ids_and_shape(path)
        if event_id in ids:
            raise RecurringAppendPlanError("Duplicate event ID")
        current = read_recurring_change_events(
            path, entries=stable, baseline_date=baseline_date,
        )
        if len(current) != len(ids):
            raise RecurringAppendPlanError("Event ledger count changed")
        if any(
            event.base_entry_id == change.base_entry_id
            and event.effective_date >= change.effective_date
            for event in current
        ):
            raise RecurringAppendPlanError(
                "Append must be later than prior events for this assignment"
            )
        preview = preview_recurring_change_from_excel(
            path, baseline_date=baseline_date, change=change,
            expected_entry=expected_entry,
            provider_max_daily_slots=provider_max_daily_slots,
            current_changes=current, require_stable_identity=True,
        )
        if preview.workbook_sha256 != before:
            raise RecurringAppendPlanError("Workbook changed during preview")
        row_values = (
            event_id, change.base_entry_id, change.patient_id,
            change.effective_date.isoformat(), change.new_therapist_id,
            change.new_start_time.strftime("%H:%M") if change.new_start_time else None,
            change.new_day_pattern, change.reason,
        )
        return RecurringAppendPlan(
            before, event_id, row_values, len(current), len(stable),
        )
    finally:
        if not path.is_file() or file_digest(path) != before:
            raise RecurringAppendPlanError("Workbook changed during read-only event plan")
