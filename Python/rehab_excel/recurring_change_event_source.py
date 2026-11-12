"""Read-only, append-only *proposed* permanent recurring change event schema.

No workbook saves, native Excel COM, clinical changes or patient data logging.
The optional sheet must be explicitly installed in a fresh synthetic copy
before any operational transaction is designed. Missing sheet means no events,
but a malformed present sheet always fails closed.

The stable AssignmentID is immutable; history is a sequence of dated events.
Each event changes one or more fields, leaving previous service dates intact.
"""
from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook

from rehab_core.models import BaseScheduleEntry
from rehab_core.recurring_therapy_changes import (
    RecurringTherapyChange, recurring_entries_on_date,
)
from .stable_recurring_ids import is_stable_assignment_id


SHEET_NAME = "RECURRING_CHANGE_EVENTS"
HEADERS = (
    "EventID", "AssignmentID", "PatientID", "EffectiveDate",
    "NewTherapist", "NewStartTime", "NewDayPattern", "Reason",
)


class RecurringEventSourceError(ValueError):
    """Stored change events are incomplete, ambiguous or inconsistent."""


def _blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _text(value):
    return None if _blank(value) else str(value).strip()


def _read_date(value):
    if isinstance(value, datetime):
        if value.time() != time():
            raise RecurringEventSourceError("Effective date cannot contain time")
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value.strip())
        except ValueError:
            pass
    raise RecurringEventSourceError("Invalid effective date")


def _read_time(value):
    if _blank(value):
        return None
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        if value.second or value.microsecond:
            raise RecurringEventSourceError("Start time must use whole minutes")
        return value
    if isinstance(value, str):
        try:
            value = time.fromisoformat(value.strip())
            if value.second or value.microsecond:
                raise ValueError("fractional minute")
            return value
        except ValueError:
            pass
    raise RecurringEventSourceError("Invalid changed start time")


def parse_recurring_change_event_rows(
    rows,
    *,
    entries: tuple[BaseScheduleEntry, ...],
    baseline_date: date,
) -> tuple[RecurringTherapyChange, ...]:
    """Validate every event including uniqueness, identities and chronology.

    EventID is a permanent opaque stable token. Existing events cannot be
    silently rewritten; write-side append-only semantics are not yet enabled.
    """
    by_id = {entry.base_entry_id: entry for entry in entries}
    if len(by_id) != len(entries):
        raise RecurringEventSourceError("Duplicate recurring assignment identity")
    if any(not is_stable_assignment_id(k) for k in by_id):
        raise RecurringEventSourceError("Events require exclusively stable assignment IDs")
    seen_events: set[str] = set()
    seen_dates: set[tuple[str, date]] = set()
    latest_by_assignment: dict[str, date] = {}
    changes: list[RecurringTherapyChange] = []
    for values in rows:
        values = tuple(values)
        if len(values) != len(HEADERS):
            raise RecurringEventSourceError("Unexpected event row width")
        if all(_blank(value) for value in values):
            continue
        event_id, assignment, patient, effective, provider, clock, days, reason = values
        event_id = _text(event_id)
        assignment = _text(assignment)
        patient = _text(patient)
        if not event_id or not event_id.startswith("change:") or len(event_id) < 16:
            raise RecurringEventSourceError("Invalid immutable change event ID")
        if event_id in seen_events:
            raise RecurringEventSourceError("Duplicate change event ID")
        seen_events.add(event_id)
        if not assignment or not is_stable_assignment_id(assignment):
            raise RecurringEventSourceError("Event references nonstable assignment")
        entry = by_id.get(assignment)
        if entry is None or entry.patient_id != patient:
            raise RecurringEventSourceError("Change event assignment/patient mismatch")
        event_date = _read_date(effective)
        if event_date < baseline_date:
            raise RecurringEventSourceError("Change event predates programme baseline")
        key = (assignment, event_date)
        if key in seen_dates:
            raise RecurringEventSourceError("Same-day change event collision")
        seen_dates.add(key)
        previous = latest_by_assignment.get(assignment)
        if previous is not None and event_date <= previous:
            raise RecurringEventSourceError("Change event order is not append-only chronological")
        latest_by_assignment[assignment] = event_date
        proposed_provider = _text(provider)
        proposed_clock = _read_time(clock)
        proposed_days = _text(days)
        if not any((proposed_provider, proposed_clock, proposed_days)):
            raise RecurringEventSourceError("Change event has no changed fields")
        try:
            change = RecurringTherapyChange(
                base_entry_id=assignment,
                patient_id=patient,
                effective_date=event_date,
                new_therapist_id=proposed_provider,
                new_start_time=proposed_clock,
                new_day_pattern=proposed_days,
                reason=_text(reason) or "",
            )
        except ValueError as exc:
            raise RecurringEventSourceError("Invalid recurring event semantics") from exc
        changes.append(change)
    # The projection path also proves no duplicate base entry/date events.
    if changes:
        recurring_entries_on_date(
            entries, baseline_date=baseline_date,
            target_date=max(c.effective_date for c in changes),
            changes=changes,
        )
    return tuple(changes)


def read_recurring_change_events(
    workbook_path: str | Path,
    *,
    entries: tuple[BaseScheduleEntry, ...],
    baseline_date: date,
) -> tuple[RecurringTherapyChange, ...]:
    """Read optional event storage but never save it."""
    path = Path(workbook_path)
    workbook = load_workbook(
        path, read_only=True, data_only=True,
        keep_vba=path.suffix.lower() == ".xlsm",
    )
    try:
        if SHEET_NAME not in workbook.sheetnames:
            return ()
        ws = workbook[SHEET_NAME]
        headers = tuple(ws.cell(1, i).value for i in range(1, len(HEADERS) + 1))
        if headers != HEADERS:
            raise RecurringEventSourceError("Recurring change event sheet headers differ")
        # Reject hidden extra columns; a schema change must be intentional.
        if ws.max_column > len(HEADERS):
            raise RecurringEventSourceError("Unexpected change event columns")
        return parse_recurring_change_event_rows(
            ws.iter_rows(min_row=2, max_col=len(HEADERS), values_only=True),
            entries=entries, baseline_date=baseline_date,
        )
    finally:
        workbook.close()
