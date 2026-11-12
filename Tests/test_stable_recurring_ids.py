"""Read-only assignment-ID contract on entirely synthetic XLSX fixtures."""
from datetime import time

import pytest
from openpyxl import Workbook

from rehab_core.models import BaseScheduleEntry
from rehab_excel.stable_recurring_ids import (
    StableRecurringIdentityError,
    bind_stable_assignment_ids,
    is_stable_assignment_id,
)

P_ID = "assignment:ABCDEF12345678"
W_ID = "assignment:WEEKLY12345678"
O_ID = "assignment:OUTPAT12345678"


def _entry(source, patient, treatment, hour):
    return BaseScheduleEntry(source, patient, treatment, time(hour, 15),
                             "Δε", "DEMO-T")


def _workbook(tmp_path, *, planner=True, weekly=True, outpatient=True):
    path = tmp_path / "synthetic_identity.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "RECURRING_ASSIGNMENT_IDS"
    ws.append(["AssignmentID", "PatientID", "Treatment"])
    if planner:
        ws.append([P_ID, "P-1", "ΦΘ"])
    ws = wb.create_sheet("PROGRAM_WEEKLY_SLOTS")
    ws.append(["PatientID", "Treatment", "Weekday", "StartTime",
               "Therapist", "SourceCell", "AssignmentID"])
    if weekly:
        ws.append(["P-2", "ΦΘ", "ΔΕ", "09:15", "DEMO-T", "C4", W_ID])
    ws = wb.create_sheet("OUTPATIENT_SCHEDULE")
    ws.append(["PatientID", "Ασθενής", "Θεραπεία",
               "Ώρα", "Ημέρες", "Θεραπευτής", "AssignmentID"])
    if outpatient:
        ws.append(["O-1", "SYNTHETIC PERSON", "ΦΘ", "11:15", "Δε", "DEMO-T", O_ID])
    wb.save(path)
    wb.close()
    return path


def test_all_three_sources_bind_without_using_row_number_as_identity(tmp_path):
    path = _workbook(tmp_path)
    entries = (
        _entry("planner:12:ΦΘ", "P-1", "ΦΘ", 10),
        _entry("weekly:2", "P-2", "ΦΘ", 9),
        _entry("outpatient:2", "O-1", "ΦΘ", 11),
    )
    before = path.read_bytes()
    resolved = bind_stable_assignment_ids(path, entries)
    assert [x.base_entry_id for x in resolved] == [P_ID, W_ID, O_ID]
    assert path.read_bytes() == before


def test_planner_id_survives_moved_excel_row(tmp_path):
    path = _workbook(tmp_path)
    for row in (2, 89):
        result = bind_stable_assignment_ids(
            path, (_entry(f"planner:{row}:ΦΘ", "P-1", "ΦΘ", 10),)
        )
        assert result[0].base_entry_id == P_ID


def test_missing_mapping_fails_closed(tmp_path):
    path = _workbook(tmp_path, planner=False)
    with pytest.raises(StableRecurringIdentityError, match="lacks immutable"):
        bind_stable_assignment_ids(
            path, (_entry("planner:2:ΦΘ", "P-1", "ΦΘ", 10),)
        )


def test_duplicate_assignment_id_fails_closed(tmp_path):
    path = _workbook(tmp_path)
    from openpyxl import load_workbook
    book = load_workbook(path)
    book["PROGRAM_WEEKLY_SLOTS"]["G2"] = P_ID
    book.save(path)
    book.close()
    with pytest.raises(StableRecurringIdentityError, match="Duplicate immutable"):
        bind_stable_assignment_ids(path, (
            _entry("planner:2:ΦΘ", "P-1", "ΦΘ", 10),
            _entry("weekly:2", "P-2", "ΦΘ", 9),
        ))


def test_shifted_weekly_row_is_not_silently_mapped_to_another_patient(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(StableRecurringIdentityError, match="Source row changed"):
        bind_stable_assignment_ids(path, (
            _entry("weekly:2", "SOMEONE-ELSE", "ΦΘ", 9),
        ))


def test_duplicate_planner_patient_treatment_is_ambiguous(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(StableRecurringIdentityError, match="Duplicate planner"):
        bind_stable_assignment_ids(path, (
            _entry("planner:2:ΦΘ", "P-1", "ΦΘ", 10),
            _entry("planner:7:ΦΘ", "P-1", "ΦΘ", 11),
        ))


def test_partial_legacy_view_allowed_only_explicitly(tmp_path):
    path = _workbook(tmp_path, planner=False)
    entry = _entry("planner:2:ΦΘ", "P-1", "ΦΘ", 10)
    assert bind_stable_assignment_ids(
        path, (entry,), require_complete=False
    ) == (entry,)


@pytest.mark.parametrize("identity", [
    "weekly:2", "outpatient:2", "assignment:short",
    "assignment:with spaces and symbols", "assignment:", "",
])
def test_row_based_or_malformed_identity_is_never_persistent(identity):
    assert not is_stable_assignment_id(identity)


def test_stable_id_format():
    assert is_stable_assignment_id(P_ID)


def test_duplicate_registry_id_is_rejected_even_if_only_one_key_requested(tmp_path):
    path = _workbook(tmp_path)
    from openpyxl import load_workbook
    book = load_workbook(path)
    book["RECURRING_ASSIGNMENT_IDS"].append([P_ID, "P-UNUSED", "ΦΘ"])
    book.save(path)
    book.close()
    with pytest.raises(StableRecurringIdentityError, match="Duplicate immutable planner"):
        bind_stable_assignment_ids(path, (
            _entry("planner:2:ΦΘ", "P-1", "ΦΘ", 10),
        ))


def test_embedded_id_survives_row_insertion_when_entry_row_tracks_move(tmp_path):
    path = _workbook(tmp_path)
    from openpyxl import load_workbook
    book = load_workbook(path)
    ws = book["PROGRAM_WEEKLY_SLOTS"]
    ws.insert_rows(2)
    book.save(path)
    book.close()
    moved_entry = _entry("weekly:3", "P-2", "ΦΘ", 9)
    result = bind_stable_assignment_ids(path, (moved_entry,))
    assert result[0].base_entry_id == W_ID
