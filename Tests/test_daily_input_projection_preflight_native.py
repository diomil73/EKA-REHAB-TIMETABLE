"""Synthetic tests for the one-call non-authorizing native projection path."""
from datetime import date, time
from hashlib import sha256

import pytest

import rehab_excel.daily_input_projection_preflight_native as native
from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import BaseScheduleEntry, Patient, Session
from rehab_excel.daily_input_input_ledger_native import InputLedgerNativeReadback
from rehab_excel.daily_input_projection_native import DailyInputProjectionNativeReadback
from rehab_excel.daily_input_projection_preflight import DailyInputProjectionPreflight
from rehab_excel.daily_input_reader import DailyInputReadResult, DailyInputSourceChange
from rehab_excel.daily_input_therapist_projection_expected import TherapistDailyExpectedGrid
from rehab_excel.daily_input_therapist_projection_native import TherapistDailyCellExpectation
from rehab_excel.reader import WorkbookSettings


SERVICE_DATE = date(2026, 11, 23)
BASELINE_DATE = date(2026, 1, 5)


def _path(tmp_path):
    path = tmp_path / "full-native-projection.xlsm"
    path.write_bytes(b"synthetic disposable workbook")
    return path


def _preflight(path):
    digest = sha256(path.read_bytes()).hexdigest()
    patient = Patient("P1", "FICTIONAL PATIENT", room="A1", status="Active")
    entry = BaseScheduleEntry(
        "assignment:00000001", "P1", "ΦΘ", time(8, 30), "Δε", "THERAPIST 1",
    )
    session = Session(
        "assignment:00000001@2026-11-23", "P1", "THERAPIST 1",
        SERVICE_DATE, time(8, 30), "ΦΘ",
    )
    state = DailySessionState(
        session_id=session.session_id, patient_id="P1", session_date=SERVICE_DATE,
        status=DailySessionStatus.THERAPIST_ABSENT,
        original_therapist_id="THERAPIST 1", original_time=time(8, 30),
        effective_therapist_id=None, effective_time=None,
    )
    daily = DailyInputReadResult(
        target_date=SERVICE_DATE, absences=(), therapist_rows_used=1,
        patient_rows_used=0,
        source_changes=(DailyInputSourceChange(
            "change:0001", 7, "therapist_absence", "THERAPIST 1",
            "THERAPIST 1", "fictional", True,
        ),),
    )
    headers = ("THERAPIST 1",) + ("",) * 17
    cells = tuple(
        TherapistDailyCellExpectation(f"{chr(ord('B') + col)}{row}", None)
        for row in (*range(2, 9), *range(12, 19))
        for col in range(9)
    )
    grid = TherapistDailyExpectedGrid(
        service_date=SERVICE_DATE, source_workbook_sha256=digest,
        provider_headers=headers, cells=cells,
        expected_assignments=(), occupied_cells=0,
    )
    settings = WorkbookSettings(
        therapist_names=("THERAPIST 1",), standard_timeslots=(time(8, 30),),
        reclined_timeslots=(), day_patterns=("Δε",), therapies=("ΦΘ",),
        patient_statuses=("Active",), yes_no_values=("ΝΑΙ", "ΟΧΙ"), rooms=("A1",),
    )
    return DailyInputProjectionPreflight(
        service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
        source_workbook_sha256=digest,
        patients=(patient,), settings=settings,
        stable_base_entries=(entry,), effective_base_entries=(entry,),
        recurring_changes=(), patient_exits=(), sessions=(session,),
        daily_input=daily, replacements=(), states=(state,), students=(),
        expected_projection=(), expected_therapist_grid=grid,
        approved_shared_memberships=(),
    )


def _results(path):
    digest = sha256(path.read_bytes()).hexdigest()
    ledger = InputLedgerNativeReadback(
        service_date=SERVICE_DATE, workbook_sha256=digest,
        source_change_count=1, active_ledger_count=1, native_excel_version="16.0",
    )
    projections = DailyInputProjectionNativeReadback(
        workbook_sha256=digest, service_date=SERVICE_DATE,
        master_projected_rows=0, therapist_grid_cells_checked=126,
        therapist_occupied_cells=0, expected_therapist_assignments=0,
        input_ledger_source_changes=1, input_ledger_active_rows=1,
        master_excel_version="16.0", therapist_excel_version="16.0",
    )
    return ledger, projections


def test_full_native_projection_wrapper_runs_all_three_readbacks(tmp_path, monkeypatch):
    path = _path(tmp_path)
    pre = _preflight(path)
    ledger, projections = _results(path)
    calls = []
    input_factory = lambda: "input-factory"
    master_factory = lambda: "master-factory"
    therapist_factory = lambda: "therapist-factory"

    def input_reader(source, **kwargs):
        calls.append(("input", source, kwargs))
        return ledger

    def projection_reader(source, **kwargs):
        calls.append(("projection", source, kwargs))
        return projections

    monkeypatch.setattr(native, "verify_input_and_ledger_native_readonly", input_reader)
    monkeypatch.setattr(native, "verify_daily_input_projection_native_readonly", projection_reader)

    result = native.verify_daily_input_projection_preflight_native_readonly(
        path, preflight=pre,
        input_ledger_excel_factory=input_factory,
        master_excel_factory=master_factory,
        therapist_excel_factory=therapist_factory,
    )

    assert result.source_workbook_sha256 == pre.source_workbook_sha256
    assert result.patient_count == 1
    assert result.base_assignment_count == 1
    assert result.session_count == 1
    assert result.daily_change_count == 1
    assert result.native_readback.therapist_grid_cells_checked == 126
    assert calls[0][0] == "input"
    assert calls[0][2]["validated_changes"] == pre.validated_changes
    assert calls[0][2]["patients"] == pre.patients
    assert calls[0][2]["sessions"] == pre.sessions
    assert calls[0][2]["therapist_names"] == pre.settings.therapist_names
    assert calls[0][2]["excel_factory"] is input_factory
    assert calls[1][0] == "projection"
    assert calls[1][2]["input_ledger_readback"] is ledger
    assert calls[1][2]["expected_projection"] == pre.expected_projection
    assert calls[1][2]["expected_projection_source_sha256"] == pre.source_workbook_sha256
    assert calls[1][2]["expected_therapist_grid"] is pre.expected_therapist_grid
    assert calls[1][2]["master_excel_factory"] is master_factory
    assert calls[1][2]["therapist_excel_factory"] is therapist_factory


def test_full_native_projection_wrapper_rejects_preflight_hash_drift(tmp_path, monkeypatch):
    path = _path(tmp_path)
    pre = _preflight(path)
    monkeypatch.setattr(native, "verify_input_and_ledger_native_readonly",
                        lambda *a, **k: pytest.fail("must stop before Excel"))
    path.write_bytes(path.read_bytes() + b"changed")
    with pytest.raises(native.DailyInputProjectionPreflightNativeError):
        native.verify_daily_input_projection_preflight_native_readonly(path, preflight=pre)


def test_full_native_projection_wrapper_rejects_input_ledger_binding_drift(tmp_path, monkeypatch):
    path = _path(tmp_path)
    pre = _preflight(path)
    ledger, projections = _results(path)
    wrong = InputLedgerNativeReadback(
        service_date=SERVICE_DATE, workbook_sha256="0" * 64,
        source_change_count=1, active_ledger_count=1, native_excel_version="16.0",
    )
    monkeypatch.setattr(native, "verify_input_and_ledger_native_readonly", lambda *a, **k: wrong)
    monkeypatch.setattr(native, "verify_daily_input_projection_native_readonly",
                        lambda *a, **k: pytest.fail("must stop before projection readback"))
    with pytest.raises(native.DailyInputProjectionPreflightNativeError):
        native.verify_daily_input_projection_preflight_native_readonly(path, preflight=pre)


def test_full_native_projection_wrapper_rejects_workbook_mutation_after_readbacks(tmp_path, monkeypatch):
    path = _path(tmp_path)
    pre = _preflight(path)
    ledger, projections = _results(path)
    monkeypatch.setattr(native, "verify_input_and_ledger_native_readonly", lambda *a, **k: ledger)

    def mutate(source, **kwargs):
        source.write_bytes(source.read_bytes() + b"tamper")
        return projections

    monkeypatch.setattr(native, "verify_daily_input_projection_native_readonly", mutate)
    with pytest.raises(native.DailyInputProjectionPreflightNativeError):
        native.verify_daily_input_projection_preflight_native_readonly(path, preflight=pre)


def test_full_native_projection_wrapper_rejects_ledger_count_drift(tmp_path, monkeypatch):
    path = _path(tmp_path)
    pre = _preflight(path)
    wrong = InputLedgerNativeReadback(
        service_date=SERVICE_DATE,
        workbook_sha256=pre.source_workbook_sha256,
        source_change_count=2,
        active_ledger_count=2,
        native_excel_version="16.0",
    )
    monkeypatch.setattr(native, "verify_input_and_ledger_native_readonly", lambda *a, **k: wrong)
    monkeypatch.setattr(native, "verify_daily_input_projection_native_readonly",
                        lambda *a, **k: pytest.fail("must stop before projection readback"))
    with pytest.raises(native.DailyInputProjectionPreflightNativeError):
        native.verify_daily_input_projection_preflight_native_readonly(path, preflight=pre)
