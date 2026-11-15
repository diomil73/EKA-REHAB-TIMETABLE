"""Synthetic DEMO: discharge/program-end read-only day-effective behaviour."""
from datetime import date, datetime, time

import pytest
from openpyxl import Workbook, load_workbook

from rehab_core.lifecycle import PatientExitKind
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule
from rehab_excel.patient_exit_source import (
    HEADERS, PatientExitSourceError, preview_patient_exit_from_workbook,
    read_effective_schedule_for_date, read_patient_exits,
    verify_saved_patient_exit,
)
from rehab_excel.patient_registry_source import read_patient_registry


def _workbook(tmp_path):
    path = tmp_path / "DEMO_LIFECYCLE_SYNTHETIC.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "PATIENTS"
    ws.append([
        "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση",
        "PatientType",
    ])
    ws.append(["DEMO-P1", "A1", "DEMO INPATIENT", "Ο", "Ενεργός", "Εσωτερικός"])
    ws.append(["DEMO-P2", None, "DEMO OUTPATIENT", None, "Ενεργός", "Εξωτερικός"])
    ws.append(["DEMO-P3", "A2", "DEMO STAY", "Ο", "Ενεργός", "Εσωτερικός"])
    planner = wb.create_sheet("PATIENT_PLANNER")
    planner.append([
        "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση",
        "ΦΘ_Ώρα", "ΦΘ_Ημέρες", "ΦΘ_Θεραπευτής",
    ])
    planner.append([
        "DEMO-P1", "A1", "DEMO INPATIENT", "Ο", "Ενεργός",
        "09:00", "Καθ/να", "DEMO-T1",
    ])
    planner.append(["DEMO-P2", None, "DEMO OUTPATIENT", None, "Ενεργός"])
    planner.append([
        "DEMO-P3", "A2", "DEMO STAY", "Ο", "Ενεργός",
        "10:00", "Καθ/να", "DEMO-T2",
    ])
    outpatient = wb.create_sheet("OUTPATIENT_SCHEDULE")
    outpatient.append([
        "PatientID", "Ασθενής", "Θεραπεία", "Ώρα", "Ημέρες", "Θεραπευτής",
    ])
    outpatient.append([
        "DEMO-P2", "DEMO OUTPATIENT", "ΦΘ", "11:00", "Καθ/να", "DEMO-T3",
    ])
    exits = wb.create_sheet("PATIENT_EXITS")
    exits.append(HEADERS)
    exits.append(["DEMO-P1", "inpatient_discharge", date(2026, 11, 16), "DEMO ONLY"])
    exits.append(["DEMO-P2", "outpatient_program_end", date(2026, 11, 17), "DEMO ONLY"])
    wb.save(path)
    wb.close()
    return path


def _edit(path, operation):
    wb = load_workbook(path)
    operation(wb)
    wb.save(path)
    wb.close()


def test_date_effective_exits_filter_future_but_not_past_bookings(tmp_path):
    path = _workbook(tmp_path)
    original = path.read_bytes()
    assert len(read_unified_base_schedule(path)) == 3
    monday = read_effective_schedule_for_date(path, date(2026, 11, 16))
    tuesday = read_effective_schedule_for_date(path, date(2026, 11, 17))
    wednesday = read_effective_schedule_for_date(path, date(2026, 11, 18))
    assert {s.patient_id for s in monday} == {"DEMO-P1", "DEMO-P2", "DEMO-P3"}
    assert {s.patient_id for s in tuesday} == {"DEMO-P2", "DEMO-P3"}
    assert {s.patient_id for s in wednesday} == {"DEMO-P3"}
    assert len(read_patient_registry(path)) == 3
    assert len(read_unified_base_schedule(path)) == 3
    assert path.read_bytes() == original


def test_preview_existing_exit_is_blocked_and_no_patient_is_removed(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(PatientExitSourceError, match="already has an effective exit"):
        preview_patient_exit_from_workbook(
            path, patient_id="DEMO-P1", kind=PatientExitKind.INPATIENT_DISCHARGE,
            last_active_date=date(2026, 11, 19),
        )
    assert len(read_patient_registry(path)) == 3


def test_preview_unexited_patient_identifies_recurring_slots(tmp_path):
    path = _workbook(tmp_path)
    result = preview_patient_exit_from_workbook(
        path, patient_id="DEMO-P3", kind=PatientExitKind.INPATIENT_DISCHARGE,
        last_active_date=date(2026, 11, 18),
    )
    assert result.patient_exit.patient_id == "DEMO-P3"
    assert len(result.affected_entries) == 1
    assert result.future_session_ids == ()


def test_mismatched_patient_type_blocks_instead_of_inventing_an_absence(tmp_path):
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENT_EXITS"].cell(3, 2, "inpatient_discharge"))
    with pytest.raises(PatientExitSourceError, match="patient type"):
        read_patient_exits(path)


def test_unknown_or_duplicate_patient_id_blocks_entire_source(tmp_path):
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENT_EXITS"].cell(3, 1, "UNREGISTERED"))
    with pytest.raises(PatientExitSourceError, match="unknown or duplicate"):
        read_patient_exits(path)
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENT_EXITS"].cell(4, 1, "DEMO-P1"))
    with pytest.raises(PatientExitSourceError, match="unknown or duplicate"):
        read_patient_exits(path)


def test_unknown_kind_and_bad_date_fail_closed(tmp_path):
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENT_EXITS"].cell(2, 2, "PATIENT_ABSENT"))
    with pytest.raises(PatientExitSourceError, match="unsupported exit kind"):
        read_patient_exits(path)
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENT_EXITS"].cell(2, 3, "INVALID"))
    with pytest.raises(PatientExitSourceError, match="date"):
        read_patient_exits(path)


def test_native_excel_date_serial_and_datetime_are_supported(tmp_path):
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENT_EXITS"].cell(2, 3, 46342))
    result = read_patient_exits(path)
    assert isinstance(result[0].last_active_date, date)
    assert all(isinstance(x.last_active_date, date) for x in result)
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENT_EXITS"].cell(2, 3, datetime(2026, 11, 16, 9)))
    assert read_patient_exits(path)[0].last_active_date == date(2026, 11, 16)


def test_optional_sheet_is_backward_compatible_but_wrong_schema_is_not(tmp_path):
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb.remove(wb["PATIENT_EXITS"]))
    assert read_patient_exits(path) == ()
    assert len(read_effective_schedule_for_date(path, date(2026, 11, 20))) == 3
    _edit(path, lambda wb: wb.create_sheet("PATIENT_EXITS").append(["BadHeader"]))
    with pytest.raises(PatientExitSourceError, match="schema"):
        read_patient_exits(path)


def test_source_dates_do_not_generate_weekend_sessions(tmp_path):
    from rehab_core.base_schedule import materialize_sessions_for_date
    path = _workbook(tmp_path)
    entries = read_effective_schedule_for_date(path, date(2026, 11, 21))
    assert materialize_sessions_for_date(entries, date(2026, 11, 21)) == []


def test_bad_service_date_blocks_without_reading_workbook(tmp_path):
    path = _workbook(tmp_path)
    with pytest.raises(PatientExitSourceError, match="valid service date"):
        read_effective_schedule_for_date(path, "2026-11-18")


def test_duplicate_registry_patient_id_blocks_lifecycle_source(tmp_path):
    path = _workbook(tmp_path)
    _edit(path, lambda wb: wb["PATIENTS"].cell(4, 1, "DEMO-P1"))
    with pytest.raises(PatientExitSourceError, match="duplicate patient identities"):
        read_patient_exits(path)


def test_independent_saved_exit_readback_preserves_weekly_source(tmp_path):
    path = _workbook(tmp_path)
    original = path.read_bytes()
    count = verify_saved_patient_exit(
        path, patient_id="DEMO-P1",
        kind=PatientExitKind.INPATIENT_DISCHARGE,
        last_active_date=date(2026, 11, 16),
    )
    assert count == 1
    assert path.read_bytes() == original


def test_saved_exit_readback_rejects_wrong_kind_or_date(tmp_path):
    path = _workbook(tmp_path)
    for kind, last_day in (
        (PatientExitKind.OUTPATIENT_PROGRAM_END, date(2026, 11, 16)),
        (PatientExitKind.INPATIENT_DISCHARGE, date(2026, 11, 17)),
    ):
        with pytest.raises(PatientExitSourceError, match="differs"):
            verify_saved_patient_exit(
                path, patient_id="DEMO-P1",
                kind=kind, last_active_date=last_day,
            )
