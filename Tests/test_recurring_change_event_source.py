"""Synthetic read-only permanent recurring event history regression tests."""
from datetime import date, time

import pytest
from openpyxl import Workbook

from rehab_core.models import BaseScheduleEntry
from rehab_core.recurring_therapy_changes import recurring_entries_on_date
from rehab_excel.recurring_change_event_source import (
    HEADERS, RecurringEventSourceError,
    parse_recurring_change_event_rows, read_recurring_change_events,
)

BASELINE = date(2026, 11, 1)
ASSIGNMENT = "assignment:SYNTHETIC000001"
PATIENT = "SYN-P"
BASE = (BaseScheduleEntry(
    ASSIGNMENT, PATIENT, "ΦΘ", time(9, 15), "Δε-Πε", "T-OLD"
),)


def row(event_id="change:SYNTHETIC00001",
        effective=date(2026, 11, 16), therapist="T-NEW",
        clock=None, days=None, assignment=ASSIGNMENT, patient=PATIENT):
    return (event_id, assignment, patient, effective, therapist,
            clock, days, "Synthetic approved change")


def test_effective_date_preserves_previous_calendar():
    change = parse_recurring_change_event_rows([row()], entries=BASE,
                                               baseline_date=BASELINE)
    older = recurring_entries_on_date(BASE, baseline_date=BASELINE,
                                       target_date=date(2026, 11, 9), changes=change)
    newer = recurring_entries_on_date(BASE, baseline_date=BASELINE,
                                       target_date=date(2026, 11, 16), changes=change)
    assert older[0].therapist_id == "T-OLD"
    assert newer[0].therapist_id == "T-NEW"
    assert BASE[0].therapist_id == "T-OLD"


def test_multiple_ordered_changes_apply_by_effective_date():
    rows = [row(), row(event_id="change:SYNTHETIC00002",
                       effective=date(2026, 12, 1),
                       therapist=None, clock="10:15")]
    events = parse_recurring_change_event_rows(rows, entries=BASE,
                                               baseline_date=BASELINE)
    projected = recurring_entries_on_date(
        BASE, baseline_date=BASELINE,
        target_date=date(2026, 12, 3), changes=events,
    )
    assert projected[0].therapist_id == "T-NEW"
    assert projected[0].start_time == time(10, 15)


@pytest.mark.parametrize("rows, issue", [
    ([row(), row()], "Duplicate change event ID"),
    ([row(), row(event_id="change:DIFFERENT00001")], "Same-day"),
    ([row(assignment="weekly:2")], "nonstable"),
    ([row(assignment="assignment:UNKNOWN123456")], "mismatch"),
    ([row(patient="WRONG")], "mismatch"),
    ([row(effective=date(2026, 10, 31))], "predates"),
    ([row(therapist=None, clock=None, days=None)], "no changed fields"),
    ([row(clock="12:27:35")], "whole minutes"),
    ([row(event_id="bad")], "event ID"),
])
def test_invalid_or_ambiguous_event_rejected(rows, issue):
    with pytest.raises(RecurringEventSourceError, match=issue):
        parse_recurring_change_event_rows(rows, entries=BASE,
                                          baseline_date=BASELINE)


def test_retroactive_insert_violating_append_order_rejected():
    rows = [
        row(effective=date(2026, 12, 1)),
        row(event_id="change:SYNTHETIC00002",
            effective=date(2026, 11, 19), therapist="T-THIRD"),
    ]
    with pytest.raises(RecurringEventSourceError, match="chronological"):
        parse_recurring_change_event_rows(rows, entries=BASE,
                                          baseline_date=BASELINE)


def test_optional_sheet_absent_is_backwards_compatible_and_no_write(tmp_path):
    path = tmp_path / "synthetic.xlsx"
    book = Workbook()
    book.save(path)
    book.close()
    before = path.read_bytes()
    assert read_recurring_change_events(path, entries=BASE,
                                        baseline_date=BASELINE) == ()
    assert path.read_bytes() == before


def test_workbook_readback_event_does_not_modify_file(tmp_path):
    path = tmp_path / "synthetic.xlsx"
    book = Workbook()
    ws = book.active
    ws.title = "RECURRING_CHANGE_EVENTS"
    ws.append(HEADERS)
    ws.append(row(effective=date(2026, 11, 16)))
    book.save(path)
    book.close()
    before = path.read_bytes()
    events = read_recurring_change_events(path, entries=BASE,
                                          baseline_date=BASELINE)
    assert len(events) == 1
    assert path.read_bytes() == before


def test_incorrect_event_sheet_schema_stops(tmp_path):
    path = tmp_path / "synthetic.xlsx"
    book = Workbook()
    ws = book.active
    ws.title = "RECURRING_CHANGE_EVENTS"
    ws.append(("WrongHeader", *HEADERS[1:]))
    book.save(path)
    book.close()
    with pytest.raises(RecurringEventSourceError, match="headers differ"):
        read_recurring_change_events(path, entries=BASE, baseline_date=BASELINE)
