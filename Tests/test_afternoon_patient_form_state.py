"""Synthetic workbook tests for isolated afternoon patient form state."""
from datetime import time

import pytest
from openpyxl import Workbook

from rehab_excel.afternoon_patient_form_state import (
    AfternoonFormStateError,
    afternoon_form_rows,
    afternoon_form_status,
    read_afternoon_patient_form_state,
)
from rehab_excel.afternoon_slots_source import HEADERS, SHEET_NAME


def workbook_file(tmp_path, *, has_pm=True, rows=()):
    book = Workbook()
    patients = book.active
    patients.title = "PATIENTS"
    patients.append(("PatientID", "Room", "Name", "Infectious", "Status"))
    patients.append(("SYNTH-P1", "A", "SYNTHETIC ONE", "", "ACTIVE"))
    patients.append(("SYNTH-P2", "B", "SYNTHETIC TWO", "", "ACTIVE"))
    if has_pm:
        ws = book.create_sheet(SHEET_NAME)
        ws.append(HEADERS)
        for item in rows:
            ws.append(item)
    path = tmp_path / "SYNTH_Afternoon_Form_State.xlsx"
    book.save(path)
    book.close()
    return path


def test_legacy_without_sheet_is_not_automatically_enabled(tmp_path):
    path = workbook_file(tmp_path, has_pm=False)
    result = read_afternoon_patient_form_state(path, patient_id="SYNTH-P1")
    assert not result.enabled and result.slots == ()
    assert "δεν έχει εγκατασταθεί" in afternoon_form_status(result)


def test_header_only_sheet_starts_empty(tmp_path):
    path = workbook_file(tmp_path)
    result = read_afternoon_patient_form_state(path, patient_id="SYNTH-P1")
    assert result.enabled and not result.slots
    assert "κενή" in afternoon_form_status(result)
    assert afternoon_form_rows(result) == ()


def test_only_selected_patient_rows_returned_without_mutation(tmp_path):
    path = workbook_file(tmp_path, rows=[
        ("SYNTH-P2", "ΦΘ", "ΤΡ", time(17), "SYNTH-PT2"),
        ("SYNTH-P1", "ΦΘ", "ΔΕ", time(16), "SYNTH-PT1"),
    ])
    bytes_before = path.read_bytes()
    result = read_afternoon_patient_form_state(path, patient_id="SYNTH-P1")
    assert afternoon_form_rows(result) == (("ΔΕ", "16:00", "SYNTH-PT1"),)
    assert path.read_bytes() == bytes_before
    assert "SYNTH-P1" not in repr(result)


def test_invalid_other_patient_row_fails_entire_form_load(tmp_path):
    path = workbook_file(tmp_path, rows=[
        ("SYNTH-P1", "ΦΘ", "ΔΕ", time(16), "SYNTH-PT1"),
        ("SYNTH-UNKNOWN", "ΦΘ", "ΔΕ", time(17), "SYNTH-PT2"),
    ])
    with pytest.raises( (AfternoonFormStateError, ValueError)):
        read_afternoon_patient_form_state(path, patient_id="SYNTH-P1")


def test_unknown_patient_rejected_even_if_afternoon_sheet_empty(tmp_path):
    path = workbook_file(tmp_path)
    with pytest.raises(AfternoonFormStateError):
        read_afternoon_patient_form_state(path, patient_id="SYNTH-UNKNOWN")


def test_malformed_sheet_is_not_treated_as_blank(tmp_path):
    path = workbook_file(tmp_path)
    book = __import__("openpyxl").load_workbook(path)
    book[SHEET_NAME].cell(1, 2).value = "Other"
    book.save(path)
    book.close()
    with pytest.raises(ValueError):
        read_afternoon_patient_form_state(path, patient_id="SYNTH-P1")


def test_empty_patient_selection_is_invalid(tmp_path):
    path = workbook_file(tmp_path)
    with pytest.raises(AfternoonFormStateError):
        read_afternoon_patient_form_state(path, patient_id=" ")


def test_no_programme_label_change_in_morning_planner(tmp_path):
    path = workbook_file(tmp_path)
    assert not read_afternoon_patient_form_state(path, patient_id="SYNTH-P1").slots
    # The helper never adds or modifies MASTER, PATIENT_PLANNER or any VBA.
    book = __import__("openpyxl").load_workbook(path, read_only=True)
    try:
        assert "PATIENT_PLANNER" not in book.sheetnames
        assert "MASTER_SCHEDULE" not in book.sheetnames
    finally:
        book.close()


def test_status_counts_sessions_not_days(tmp_path):
    path = workbook_file(tmp_path, rows=[
        ("SYNTH-P1", "ΦΘ", "ΔΕ", time(16), "SYNTH-PT1"),
        ("SYNTH-P1", "ΦΘ", "ΔΕ", time(17), "SYNTH-PT1"),
    ])
    state = read_afternoon_patient_form_state(path, patient_id="SYNTH-P1")
    assert afternoon_form_status(state) == "Απογευματινή ΦΘ: 2 καταχωρισμένες συνεδρίες."
