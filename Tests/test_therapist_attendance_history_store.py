from datetime import date
from hashlib import sha256
from pathlib import Path

import pytest
from openpyxl import Workbook

from rehab_excel.finalized_day_snapshots import HEADERS as SNAPSHOT_HEADERS
from rehab_excel.therapist_attendance_history_store import (
    HEADERS, SHEET_NAME, TherapistAttendanceHistoryStoreError,
    read_therapist_attendance_history_facts,
    read_verified_therapist_attendance,
)


DAY = date(2026, 10, 29)


def fixture_workbook(path: Path, *, include_attendance=True, facts=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "FINALIZED_DAY_SNAPSHOTS"
    ws.append(SNAPSHOT_HEADERS)
    ws.append((
        "old", DAY.isoformat(), 1, "SUPERSEDED", "2026-10-29T16:00:00+02:00",
        "", 0, 4, 4, 0, 0, 0, 0, 0, 0,
    ))
    ws.append((
        "current", DAY.isoformat(), 2, "CURRENT_FINAL", "2026-10-29T17:00:00+02:00",
        "old", 0, 4, 4, 0, 0, 0, 0, 0, 0,
    ))
    if include_attendance:
        history = wb.create_sheet(SHEET_NAME)
        history.append(HEADERS)
        for row in facts:
            history.append(row)
    wb.save(path)
    wb.close()
    return path


def record(*, snapshot="current", outcome="PRESENT", reason=""):
    return (snapshot, DAY.isoformat(), "T1", outcome, "approved roster", reason)


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def test_missing_history_sheet_is_unavailable_and_never_created(tmp_path):
    path = fixture_workbook(tmp_path / "missing.xlsx", include_attendance=False)
    before = digest(path)
    with pytest.raises(TherapistAttendanceHistoryStoreError, match="is missing"):
        read_therapist_attendance_history_facts(path)
    assert digest(path) == before


def test_current_final_attendance_and_superseded_version_are_distinguished(tmp_path):
    path = fixture_workbook(tmp_path / "versions.xlsx", facts=(
        record(snapshot="old", outcome="ABSENT", reason="SICK_LEAVE"),
        record(snapshot="current", outcome="PRESENT"),
    ))
    before = digest(path)
    summary = read_verified_therapist_attendance(
        path, therapist_id="T1", expected_working_dates=(DAY,),
    )
    assert (summary.working_days, summary.present_days, summary.absent_days) == (1, 1, 0)
    assert summary.presence_ratio == 1
    assert digest(path) == before


def test_missing_current_version_is_not_a_false_zero_absence(tmp_path):
    path = fixture_workbook(tmp_path / "superseded.xlsx", facts=(
        record(snapshot="old", outcome="ABSENT", reason="SICK_LEAVE"),
    ))
    with pytest.raises(TherapistAttendanceHistoryStoreError, match="Missing verified"):
        read_verified_therapist_attendance(
            path, therapist_id="T1", expected_working_dates=(DAY,),
        )


def test_missing_snapshot_for_expected_workday_is_safety_stop(tmp_path):
    path = fixture_workbook(tmp_path / "missing-day.xlsx", facts=(record(),))
    with pytest.raises(TherapistAttendanceHistoryStoreError, match="without a CURRENT_FINAL"):
        read_verified_therapist_attendance(
            path, therapist_id="T1", expected_working_dates=(date(2026, 10, 30),),
        )


def test_duplicate_and_invalid_source_are_refused(tmp_path):
    path = fixture_workbook(tmp_path / "duplicates.xlsx", facts=(record(), record()))
    with pytest.raises(TherapistAttendanceHistoryStoreError, match="duplicate"):
        read_therapist_attendance_history_facts(path)
    path = fixture_workbook(tmp_path / "invalid.xlsx", facts=(record(outcome="ABSENT"),))
    with pytest.raises(TherapistAttendanceHistoryStoreError, match="Absence reason"):
        read_therapist_attendance_history_facts(path)


def test_empty_valid_history_still_requires_explicit_working_day_fact(tmp_path):
    path = fixture_workbook(tmp_path / "empty.xlsx", facts=())
    with pytest.raises(TherapistAttendanceHistoryStoreError, match="Missing verified"):
        read_verified_therapist_attendance(
            path, therapist_id="T1", expected_working_dates=(DAY,),
        )
