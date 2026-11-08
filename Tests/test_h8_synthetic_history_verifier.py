from datetime import date
from hashlib import sha256

import pytest
from openpyxl import Workbook, load_workbook

from tools.seed_h8_synthetic_history import _days, generate_demo_rows
from tools.verify_h8_synthetic_history import (
    DemoVerificationError,
    verify_demo_history,
)
from rehab_excel.finalized_day_snapshots import HEADERS as SNAP_HEADERS
from rehab_excel.therapist_day_history_store import HEADERS as THERAPIST_HEADERS
from rehab_excel.patient_session_history_store import HEADERS as PATIENT_HEADERS
from rehab_excel.day_history_event_store import HEADERS as EVENT_HEADERS
from rehab_excel.therapist_attendance_history_store import HEADERS as ATTENDANCE_HEADERS


HEADERS = {
    "FINALIZED_DAY_SNAPSHOTS": SNAP_HEADERS,
    "THERAPIST_DAY_HISTORY": THERAPIST_HEADERS,
    "PATIENT_SESSION_HISTORY": PATIENT_HEADERS,
    "DAY_HISTORY_EVENTS": EVENT_HEADERS,
    "THERAPIST_ATTENDANCE_HISTORY": ATTENDANCE_HEADERS,
}


def fixture_book(path):
    workbook = Workbook()
    registry = workbook.active
    registry.title = "PATIENTS"
    registry.append(("Id", "Room", "Name", "Infectious", "Status"))
    patients = tuple(f"TEST-PATIENT-{i:02d}" for i in range(36))
    for patient in patients:
        registry.append((patient, "", patient, "", "ACTIVE"))
    dates = _days(date(2026, 9, 21), date(2026, 10, 23), set())
    therapists = tuple(f"TEST-THERAPIST-{i}" for i in range(6))
    rows = generate_demo_rows(dates, therapists, patients)
    for name, headers in HEADERS.items():
        sheet = workbook.create_sheet(name)
        sheet.append(headers)
        for row in rows[name]:
            sheet.append(row)
    workbook.save(path)
    workbook.close()
    return path


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def change(path, name, row, col, value):
    book = load_workbook(path, keep_vba=True)
    book[name].cell(row, col).value = value
    book.save(path)
    book.close()


def test_synthetic_cross_sheet_verification_is_read_only(tmp_path):
    path = fixture_book(tmp_path / "H8_DEMO_TEST.xlsm")
    before = digest(path)
    result = verify_demo_history(path)
    assert result["days"] == 25
    assert result["therapist_days"] == 150
    assert result["attendance_facts"] == 150
    assert result["finalization_events"] == 25
    assert result["patients"] > 0
    assert digest(path) == before


def test_rejects_corrupt_snapshot_totals_and_does_not_modify_file(tmp_path):
    path = fixture_book(tmp_path / "H8_DEMO_TOTALS.xlsm")
    change(path, "FINALIZED_DAY_SNAPSHOTS", 2, 9, 999)
    before = digest(path)
    with pytest.raises(DemoVerificationError, match="performed count mismatch"):
        verify_demo_history(path)
    assert digest(path) == before


def test_rejects_missing_attendance_fact(tmp_path):
    path = fixture_book(tmp_path / "H8_DEMO_MISSING.xlsm")
    change(path, "THERAPIST_ATTENDANCE_HISTORY", 2, 1, "BAD-SNAPSHOT-ID")
    with pytest.raises(DemoVerificationError, match="expected 6 therapist"):
        verify_demo_history(path)


def test_rejects_therapist_workload_mismatch(tmp_path):
    path = fixture_book(tmp_path / "H8_DEMO_WORKLOAD.xlsm")
    change(path, "THERAPIST_DAY_HISTORY", 2, 8, 999)
    with pytest.raises(DemoVerificationError, match="delivered therapist totals mismatch"):
        verify_demo_history(path)


def test_rejects_missing_finalization_event(tmp_path):
    path = fixture_book(tmp_path / "H8_DEMO_EVENT.xlsm")
    change(path, "DAY_HISTORY_EVENTS", 2, 11, "WRONG-SNAPSHOT")
    with pytest.raises(DemoVerificationError, match="missing synthetic finalization event"):
        verify_demo_history(path)


def test_refuses_non_demo_filename(tmp_path):
    path = fixture_book(tmp_path / "PRODUCTION.xlsm")
    with pytest.raises(DemoVerificationError, match="Refusing a non-demo workbook"):
        verify_demo_history(path)
