from datetime import date, datetime, time, timezone

import pytest
from openpyxl import Workbook

from rehab_core.daily_state import DailySessionStatus
from rehab_core.patient_session_history import PatientSessionHistoryRow
from rehab_excel.finalized_day_snapshots import HEADERS as SNAPSHOT_HEADERS
from rehab_excel.patient_session_history_store import (
    HEADERS,
    PatientSessionHistoryStoreError,
    read_current_patient_session_history_rows,
    read_patient_session_history_rows,
    values_for_payload,
)


SERVICE_DATE = date(2026, 10, 30)


def _write_workbook(path, *, patient_rows, snapshot_rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "FINALIZED_DAY_SNAPSHOTS"
    ws.append(SNAPSHOT_HEADERS)
    for row in snapshot_rows:
        ws.append(row)

    patient_ws = wb.create_sheet("PATIENT_SESSION_HISTORY")
    patient_ws.append(HEADERS)
    for row in patient_rows:
        patient_ws.append(row)
    wb.save(path)
    wb.close()


def _snapshot(snapshot_id: str, version: int, status: str):
    return (
        snapshot_id,
        SERVICE_DATE.isoformat(),
        version,
        status,
        datetime(2026, 10, 30, 18, 0, tzinfo=timezone.utc).isoformat(),
        "" if version == 1 else "finalized:2026-10-30:v1",
        0,
        2,
        2,
        0,
        0,
        0,
        0,
        0,
        0,
    )


def _patient_row(snapshot_id: str, version: int, session_id: str, patient_id: str):
    return (
        snapshot_id,
        SERVICE_DATE.isoformat(),
        version,
        session_id,
        patient_id,
        "active",
        "T1",
        "09:00:00",
        "T1",
        "09:00:00",
        "",
        "",
        "",
        "",
    )


def test_reads_patient_session_rows_and_preserves_outcome_fields(tmp_path):
    path = tmp_path / "history.xlsx"
    row = list(_patient_row("finalized:2026-10-30:v2", 2, "S1", "P1"))
    row[5] = "cancelled"
    row[8] = ""
    row[9] = ""
    row[11] = "C1"
    row[12] = "patient_no_show"
    row[13] = "no show"
    _write_workbook(
        path,
        patient_rows=[row],
        snapshot_rows=[_snapshot("finalized:2026-10-30:v2", 2, "CURRENT_FINAL")],
    )

    rows = read_patient_session_history_rows(path)

    assert len(rows) == 1
    assert rows[0].session_id == "S1"
    assert rows[0].patient_id == "P1"
    assert rows[0].status == DailySessionStatus.CANCELLED
    assert rows[0].original_time == time(9, 0)
    assert rows[0].effective_therapist_id is None
    assert rows[0].effective_time is None
    assert rows[0].cancellation_kind == "patient_no_show"
    assert rows[0].reason == "no show"


def test_current_reader_excludes_superseded_snapshot_rows(tmp_path):
    path = tmp_path / "history.xlsx"
    _write_workbook(
        path,
        patient_rows=[
            _patient_row("finalized:2026-10-30:v1", 1, "S1", "P1"),
            _patient_row("finalized:2026-10-30:v2", 2, "S1", "P1"),
            _patient_row("finalized:2026-10-30:v2", 2, "S2", "P2"),
        ],
        snapshot_rows=[
            _snapshot("finalized:2026-10-30:v1", 1, "SUPERSEDED"),
            _snapshot("finalized:2026-10-30:v2", 2, "CURRENT_FINAL"),
        ],
    )

    rows = read_current_patient_session_history_rows(path)

    assert [(row.snapshot_version, row.session_id) for row in rows] == [(2, "S1"), (2, "S2")]


def test_current_reader_supports_patient_filter(tmp_path):
    path = tmp_path / "history.xlsx"
    _write_workbook(
        path,
        patient_rows=[
            _patient_row("finalized:2026-10-30:v2", 2, "S1", "P1"),
            _patient_row("finalized:2026-10-30:v2", 2, "S2", "P2"),
        ],
        snapshot_rows=[_snapshot("finalized:2026-10-30:v2", 2, "CURRENT_FINAL")],
    )

    rows = read_current_patient_session_history_rows(path, patient_id="p2")

    assert [row.session_id for row in rows] == ["S2"]


def test_rejects_duplicate_session_key_and_unresolved_status(tmp_path):
    duplicate_path = tmp_path / "duplicate.xlsx"
    row = _patient_row("finalized:2026-10-30:v2", 2, "S1", "P1")
    _write_workbook(
        duplicate_path,
        patient_rows=[row, row],
        snapshot_rows=[_snapshot("finalized:2026-10-30:v2", 2, "CURRENT_FINAL")],
    )
    with pytest.raises(PatientSessionHistoryStoreError, match="Duplicate patient/session history key"):
        read_patient_session_history_rows(duplicate_path)

    unresolved_path = tmp_path / "unresolved.xlsx"
    unresolved = list(row)
    unresolved[5] = "therapist_absent"
    unresolved[8] = ""
    unresolved[9] = ""
    _write_workbook(
        unresolved_path,
        patient_rows=[unresolved],
        snapshot_rows=[_snapshot("finalized:2026-10-30:v2", 2, "CURRENT_FINAL")],
    )
    with pytest.raises(PatientSessionHistoryStoreError, match="unresolved therapist absence"):
        read_patient_session_history_rows(unresolved_path)


def test_rejects_patient_row_snapshot_version_mismatch(tmp_path):
    path = tmp_path / "history.xlsx"
    _write_workbook(
        path,
        patient_rows=[_patient_row("finalized:2026-10-30:v2", 1, "S1", "P1")],
        snapshot_rows=[_snapshot("finalized:2026-10-30:v2", 2, "CURRENT_FINAL")],
    )

    with pytest.raises(PatientSessionHistoryStoreError, match="version does not match"):
        read_current_patient_session_history_rows(path)


def test_payload_order_is_stable():
    row = PatientSessionHistoryRow(
        snapshot_id="finalized:2026-10-30:v2",
        service_date=SERVICE_DATE,
        snapshot_version=2,
        session_id="S1",
        patient_id="P1",
        status=DailySessionStatus.REPLACED,
        original_therapist_id="T1",
        original_time=time(9, 0),
        effective_therapist_id="T2",
        effective_time=time(9, 30),
        replacement_id="R1",
        cancellation_id=None,
        cancellation_kind=None,
        reason="replacement",
    )

    assert values_for_payload(row) == (
        "finalized:2026-10-30:v2",
        "2026-10-30",
        2,
        "S1",
        "P1",
        "replaced",
        "T1",
        "09:00:00",
        "T2",
        "09:30:00",
        "R1",
        None,
        None,
        "replacement",
    )
