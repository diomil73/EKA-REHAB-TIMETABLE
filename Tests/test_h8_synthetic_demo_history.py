from collections import Counter
from datetime import date, timedelta
from pathlib import Path

import pytest

from rehab_core.daily_state import DailySessionStatus
from tools.seed_h8_synthetic_history import (
    ATTENDANCE_SHEET, DEMO_SOURCE, DemoHistoryError,
    _days, generate_demo_rows,
)
from tools.prepare_h8_synthetic_demo import DemoPreparationError, prepare_demo


def cohort():
    dates = _days(date(2026, 9, 21), date(2026, 10, 23), set())
    therapist_ids = tuple(f"THERAPIST-{i}" for i in range(6))
    patient_ids = tuple(f"PATIENT-{i:02d}" for i in range(36))
    return dates, therapist_ids, patient_ids


def test_seed_has_cross_week_and_month_dates_and_is_reproducible():
    args = cohort()
    a = generate_demo_rows(*args)
    b = generate_demo_rows(*args)
    assert a == b
    assert len(a["FINALIZED_DAY_SNAPSHOTS"]) >= 20
    months = {row[1][:7] for row in a["FINALIZED_DAY_SNAPSHOTS"]}
    assert months == {"2026-09", "2026-10"}
    assert len(a["THERAPIST_DAY_HISTORY"]) == 6 * len(args[0])
    assert len(a[ATTENDANCE_SHEET]) == 6 * len(args[0])
    assert {row[3] for row in a[ATTENDANCE_SHEET]} == {"PRESENT", "ABSENT"}
    assert {row[4] for row in a[ATTENDANCE_SHEET]} == {DEMO_SOURCE}


def test_snapshot_outcomes_reconcile_exactly_with_daily_patient_sessions():
    data = generate_demo_rows(*cohort())
    for snapshot in data["FINALIZED_DAY_SNAPSHOTS"]:
        snapshot_id = snapshot[0]
        rows = [r for r in data["PATIENT_SESSION_HISTORY"] if r[0] == snapshot_id]
        counts = Counter(r[5] for r in rows)
        assert len(rows) == snapshot[7]
        assert counts["active"] + counts["replaced"] == snapshot[8]
        assert counts["patient_absent"] == snapshot[9]
        assert counts["cancelled"] == snapshot[10]
        assert counts["replaced"] == snapshot[12]
        assert sum(1 for r in rows if r[12] == "treatment_suspended") == snapshot[13]
        assert len({r[3] for r in rows}) == len(rows)
        assert len({r[4] for r in rows}) == len(rows)
        assert all(r[5] != DailySessionStatus.THERAPIST_ABSENT.value for r in rows)
        assert snapshot[6] == 0


def test_therapist_rows_and_absences_reconcile_against_delivered_sessions():
    data = generate_demo_rows(*cohort())
    sessions = data["PATIENT_SESSION_HISTORY"]
    for snapshot, day, therapist, capacity, planned_slots, planned, actual_slots, active, replacement, released, suspended, approved in data["THERAPIST_DAY_HISTORY"]:
        scheduled = [r for r in sessions if r[0] == snapshot and r[6] == therapist]
        performed = [r for r in sessions if r[0] == snapshot and r[8] == therapist and r[5] in ("active", "replaced")]
        assert planned == len(scheduled) == planned_slots
        assert active == len(performed) == actual_slots
        assert capacity == 6
        assert replacement == sum(r[5] == "replaced" for r in performed)
        assert released == sum(
            r[5] == "patient_absent" or r[12] == "department_postponed"
            for r in scheduled
        )
        assert suspended == sum(r[12] == "treatment_suspended" for r in scheduled)
        assert approved == 0
    for snapshot, day, therapist, outcome, source, reason in data[ATTENDANCE_SHEET]:
        if outcome == "ABSENT":
            assert reason == "SICK_LEAVE"
            assert not any(
                r[0] == snapshot and r[8] == therapist
                and r[5] in ("active", "replaced")
                for r in sessions
            )
        else:
            assert not reason


def test_demo_generation_fails_on_weak_roster_weekends_and_duplicates():
    days, therapists, patients = cohort()
    with pytest.raises(DemoHistoryError):
        generate_demo_rows(days[:11], therapists, patients)
    with pytest.raises(DemoHistoryError):
        generate_demo_rows(days, therapists[:5], patients)
    with pytest.raises(DemoHistoryError):
        generate_demo_rows(days, therapists, patients[:23])
    with pytest.raises(DemoHistoryError):
        generate_demo_rows(days + (days[0],), therapists, patients)
    with pytest.raises(DemoHistoryError):
        generate_demo_rows(days + (date(2026, 10, 24),), therapists, patients)


def test_native_builder_fails_before_any_side_effect_if_output_already_exists(tmp_path, monkeypatch):
    source = tmp_path / "AZ.xlsm"
    output = tmp_path / "H8_DEMO_BM.xlsm"
    source.write_bytes(b"sentinel source")
    output.write_bytes(b"already in use")
    called = []
    monkeypatch.setattr(
        "tools.prepare_h8_synthetic_demo.subprocess.run",
        lambda *args, **kwargs: called.append(args),
    )
    with pytest.raises(DemoPreparationError, match="already exists"):
        prepare_demo(source, output)
    assert not called
    assert output.read_bytes() == b"already in use"
    assert source.read_bytes() == b"sentinel source"


def test_native_builder_fails_closed_with_bad_destination(tmp_path):
    source = tmp_path / "AZ.xlsm"
    source.write_bytes(b"sentinel")
    with pytest.raises(DemoPreparationError, match="_DEMO_"):
        prepare_demo(source, tmp_path / "existing_BL.xlsm")
