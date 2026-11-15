"""Synthetic regression coverage for isolated afternoon weekly source."""
from datetime import time

import pytest
from openpyxl import Workbook

from rehab_core.patient_programme_draft import prepare_afternoon_physio_drafts
from rehab_excel.afternoon_slots_source import (
    AFTERNOON_TREATMENT,
    HEADERS,
    SHEET_NAME,
    AfternoonSlot,
    AfternoonSlotsError,
    make_afternoon_rows,
    read_afternoon_slots_sheet,
    validate_afternoon_rows,
)


def sheet():
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADERS)
    return ws


def test_header_only_is_valid_and_means_no_afternoon_treatments():
    ws = sheet()
    assert read_afternoon_slots_sheet(ws, known_patient_ids=["SYNTH-P1"]) == ()


def test_blank_patient_form_never_creates_afternoon_rows():
    assert make_afternoon_rows("SYNTH-P1", prepare_afternoon_physio_drafts()) == ()


def test_explicit_afternoon_data_round_trips_from_isolated_sheet():
    drafts = prepare_afternoon_physio_drafts(
        clock="16:30", days="ΔΕ-ΤΕ", therapist="SYNTH-AFTERNOON-PT"
    )
    rows = make_afternoon_rows("SYNTH-P1", drafts)
    assert len(rows) == 2
    ws = sheet()
    for item in rows:
        ws.append((item.patient_id, item.treatment, item.weekday, item.start_time, item.therapist))
    assert read_afternoon_slots_sheet(ws, known_patient_ids=["SYNTH-P1"]) == rows


def test_afternoon_rows_are_kept_distinct_from_morning_even_same_clock():
    afternoon = AfternoonSlot("SYNTH-P1", AFTERNOON_TREATMENT, "ΔΕ", time(9), "SYNTH-PT")
    assert validate_afternoon_rows((afternoon,), known_patient_ids=["SYNTH-P1"]) == (afternoon,)


def test_same_patient_duplicate_afternoon_clock_is_blocked():
    a = AfternoonSlot("SYNTH-P1", AFTERNOON_TREATMENT, "ΔΕ", time(16), "SYNTH-PT1")
    b = AfternoonSlot("SYNTH-P1", AFTERNOON_TREATMENT, "ΔΕ", time(16), "SYNTH-PT2")
    with pytest.raises(AfternoonSlotsError):
        validate_afternoon_rows((a, b))


def test_group_therapist_clock_is_not_auto_rejected_without_group_review():
    a = AfternoonSlot("SYNTH-P1", AFTERNOON_TREATMENT, "ΔΕ", time(16), "SYNTH-PT")
    b = AfternoonSlot("SYNTH-P2", AFTERNOON_TREATMENT, "ΔΕ", time(16), "SYNTH-PT")
    assert validate_afternoon_rows((a, b)) == (a, b)


@pytest.mark.parametrize("values", [
    ("SYNTH-P1", "Εργο", "ΔΕ", time(16), "SYNTH-PT"),
    ("SYNTH-P1", "Πισίνα", "ΔΕ", time(16), "SYNTH-PT"),
    ("SYNTH-P1", "ΦΘ", "ΔΕ", time(16), None),
    ("SYNTH-P1", "ΦΘ", "", time(16), "SYNTH-PT"),
    ("SYNTH-P1", "ΦΘ", "ΔΕ", None, "SYNTH-PT"),
    ("SYNTH-P1", "ΦΘ", "ΔΕ", "16:00", "SYNTH-PT"),
])
def test_incomplete_or_non_physio_afternoon_rows_are_blocked(values):
    ws = sheet()
    ws.append(values)
    with pytest.raises(AfternoonSlotsError):
        read_afternoon_slots_sheet(ws, known_patient_ids=["SYNTH-P1"])


def test_unknown_patient_is_blocked_without_exposing_identity():
    ws = sheet()
    ws.append(("SYNTH-UNKNOWN", "ΦΘ", "ΔΕ", time(16), "SYNTH-PT"))
    with pytest.raises(AfternoonSlotsError):
        read_afternoon_slots_sheet(ws, known_patient_ids=["SYNTH-P1"])


def test_wrong_schema_is_refused():
    ws = sheet()
    ws.cell(1, 2).value = "WrongTreatmentHeader"
    with pytest.raises(AfternoonSlotsError):
        read_afternoon_slots_sheet(ws, known_patient_ids=["SYNTH-P1"])


def test_no_implicit_patient_membership():
    ws = sheet()
    with pytest.raises(AfternoonSlotsError):
        read_afternoon_slots_sheet(ws, known_patient_ids=[])


def test_invalid_non_row_object_refused():
    with pytest.raises(AfternoonSlotsError):
        validate_afternoon_rows([object()])


def test_repeated_blank_rows_not_interpreted_as_sessions():
    ws = sheet()
    ws.append((None, None, None, None, None))
    assert read_afternoon_slots_sheet(ws, known_patient_ids=["SYNTH-P1"]) == ()
