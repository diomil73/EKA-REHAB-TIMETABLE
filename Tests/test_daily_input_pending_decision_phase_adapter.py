from dataclasses import replace
from datetime import date, time
from hashlib import sha256
import inspect

import pytest

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import Session
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_pending_decision_inventory import (
    build_pending_decision_inventory, write_pending_decision_inventory,
)
from rehab_excel.daily_input_pending_decision_phase_adapter import (
    PendingDecisionPhaseReportAdapterError,
    create_pending_decision_phase_report,
)
from rehab_excel.daily_input_phase_verifier_contract import (
    PhaseVerifierNotImplementedError, validate_phase_verifier_report,
)
from rehab_excel.daily_input_projection_native import DailyInputProjectionNativeReadback
from rehab_excel.daily_input_projection_preflight import DailyInputProjectionPreflight
from rehab_excel.daily_input_projection_preflight_native import DailyInputProjectionPreflightNativeReadback
from rehab_excel.daily_input_source_native_capture import (
    DailyInputSourceNativeCapture, NativeSourceSheetResult, REQUIRED_SOURCE_SHEETS,
)
from rehab_excel.daily_input_therapist_projection_expected import TherapistDailyExpectedGrid
from rehab_excel.daily_input_reader import DailyInputReadResult, DailyInputSourceChange
from rehab_excel.reader import WorkbookSettings


DAY = date(2026, 11, 27)


def fixture(tmp_path):
    root = tmp_path / "evidence"
    root.mkdir()
    workbook = root / "synthetic-pending.xlsm"
    workbook.write_bytes(b"SYNTHETIC disposable workbook only")
    inventory_path = root / "pending-inventory.json"
    report_path = root / "pending-report.json"
    workbook_hash = sha256(workbook.read_bytes()).hexdigest()

    initial = DailyInputOperation("pending-report-demo-001", "a" * 64, DAY.isoformat())
    input_committed = advance_daily_input_operation(initial, Phase.INPUT_AND_LEDGER_COMMITTED)
    projections_verified = advance_daily_input_operation(input_committed, Phase.PROJECTIONS_VERIFIED)
    operation = advance_daily_input_operation(projections_verified, Phase.PENDING_OPERATOR_DECISION)

    session = Session("session-001", "patient-001", "therapist-001", DAY,
                      time(9, 0), "ΦΘ", True)
    state = DailySessionState(
        "session-001", "patient-001", DAY, DailySessionStatus.THERAPIST_ABSENT,
        "therapist-001", time(9, 0), None, None, reason="synthetic absence",
    )
    settings = WorkbookSettings(
        therapist_names=("therapist-001",), standard_timeslots=(time(9, 0),),
        reclined_timeslots=(), day_patterns=("ΔΕ",), therapies=("ΦΘ",),
        patient_statuses=("Active",), yes_no_values=("ΝΑΙ", "ΟΧΙ"),
        rooms=("A01",),
    )
    daily = DailyInputReadResult(
        target_date=DAY, absences=(), therapist_rows_used=1, patient_rows_used=0,
        source_changes=(DailyInputSourceChange(
            "change-001", 7, "therapist_absence", "therapist-001",
            "therapist-001", "synthetic", True,
        ),),
    )
    grid = TherapistDailyExpectedGrid(
        service_date=DAY, source_workbook_sha256=workbook_hash,
        provider_headers=("",) * 18, cells=(), expected_assignments=(),
        occupied_cells=0,
    )
    preflight = DailyInputProjectionPreflight(
        service_date=DAY, baseline_date=date(2026, 1, 5),
        source_workbook_sha256=workbook_hash, patients=(), settings=settings,
        stable_base_entries=(), effective_base_entries=(), recurring_changes=(),
        patient_exits=(), sessions=(session,), daily_input=daily, replacements=(),
        states=(state,), students=(), expected_projection=(),
        expected_therapist_grid=grid, approved_shared_memberships=(),
    )
    inventory = build_pending_decision_inventory(
        projections_verified, operation, preflight,
    )
    artifact = write_pending_decision_inventory(
        inventory, evidence_root=root, artifact_path=inventory_path,
    )
    capture = DailyInputSourceNativeCapture(
        workbook_sha256=workbook_hash, native_excel_version="16.0",
        sheets=tuple(NativeSourceSheetResult(name, 1, 1, 1)
                     for name in REQUIRED_SOURCE_SHEETS),
        identity_mirror_rows_checked=0,
    )
    projection = DailyInputProjectionNativeReadback(
        workbook_sha256=workbook_hash, service_date=DAY,
        master_projected_rows=0, therapist_grid_cells_checked=126,
        therapist_occupied_cells=0, expected_therapist_assignments=0,
        input_ledger_source_changes=1, input_ledger_active_rows=1,
        master_excel_version="16.0", therapist_excel_version="16.0",
    )
    native = DailyInputProjectionPreflightNativeReadback(
        preflight=preflight, native_readback=projection, source_capture=capture,
    )
    return (root, workbook, inventory_path, report_path, projections_verified,
            operation, preflight, inventory, artifact, native)


def create(data, **overrides):
    (root, workbook, inventory_path, report_path, previous,
     operation, preflight, _inventory, _artifact, _native) = data
    args = dict(
        previous_operation=previous, operation=operation, preflight=preflight,
        evidence_root=root, report_path=report_path, report_id="pending-report-001",
    )
    args.update(overrides)
    return create_pending_decision_phase_report(
        workbook, inventory_path, **args,
    )


def test_adapter_publishes_exact_inventory_bound_native_report_but_cannot_authorize(tmp_path, monkeypatch):
    data = fixture(tmp_path)
    native = data[-1]
    calls = []

    def fake_verify(path, **kwargs):
        calls.append((path, kwargs))
        return native

    monkeypatch.setattr(
        "rehab_excel.daily_input_pending_decision_phase_adapter.verify_daily_input_projection_preflight_native_readonly",
        fake_verify,
    )
    report = create(data)
    assert len(calls) == 1
    assert calls[0][0] == data[1].resolve()
    assert calls[0][1]["preflight"] is data[6]
    assert report.operation_id == data[5].operation_id
    assert report.journal_phase == Phase.PENDING_OPERATOR_DECISION.value
    assert report.journal_revision == data[5].revision
    assert report.source_workbook_sha256 == data[5].source_sha256
    assert report.evidence_artifact == data[2].name
    assert report.evidence_sha256 == data[8].artifact_sha256
    assert report.evidence_sha256 != data[8].artifact_sha256[:0] + data[6].source_workbook_sha256
    assert report.native_excel_version == "16.0"
    assert report.inspected_scope == (
        "pending_replacement_inventory", "operation_identity", "saved_workbook_sha256",
    )
    with pytest.raises(PhaseVerifierNotImplementedError, match="trusted verifier attestation"):
        validate_phase_verifier_report(report, data[5], previous_operation=data[4])


def test_adapter_rejects_inventory_tamper_and_preflight_drift_before_native(tmp_path, monkeypatch):
    data = fixture(tmp_path)
    calls = []
    monkeypatch.setattr(
        "rehab_excel.daily_input_pending_decision_phase_adapter.verify_daily_input_projection_preflight_native_readonly",
        lambda *a, **k: calls.append((a, k)),
    )
    payload = data[2].read_bytes().replace(b"therapist-001", b"therapist-999")
    data[2].write_bytes(payload)
    with pytest.raises(PendingDecisionPhaseReportAdapterError, match="artifact/preflight"):
        create(data)
    assert calls == []
    assert not data[3].exists()


def test_adapter_rejects_wrong_native_digest_date_scope_or_version(tmp_path, monkeypatch):
    for problem in ("hash", "date", "version", "scope"):
        case_dir = tmp_path / problem
        case_dir.mkdir()
        data = fixture(case_dir)
        native = data[-1]
        projection = native.native_readback
        capture = native.source_capture
        if problem == "hash":
            projection = replace(projection, workbook_sha256="f" * 64)
        elif problem == "date":
            projection = replace(projection, service_date=date(2026, 11, 28))
        elif problem == "version":
            capture = replace(capture, native_excel_version="15.0")
        else:
            projection = replace(projection, therapist_grid_cells_checked=125)
        bad_native = replace(native, native_readback=projection, source_capture=capture)
        monkeypatch.setattr(
            "rehab_excel.daily_input_pending_decision_phase_adapter.verify_daily_input_projection_preflight_native_readonly",
            lambda *a, _native=bad_native, **k: _native,
        )
        with pytest.raises(PendingDecisionPhaseReportAdapterError, match="scope/version mismatch"):
            create(data)
        assert not data[3].exists()


def test_adapter_rejects_workbook_mutation_during_native_verification(tmp_path, monkeypatch):
    data = fixture(tmp_path)

    def mutate(path, **kwargs):
        path.write_bytes(path.read_bytes() + b"changed")
        return data[-1]

    monkeypatch.setattr(
        "rehab_excel.daily_input_pending_decision_phase_adapter.verify_daily_input_projection_preflight_native_readonly",
        mutate,
    )
    with pytest.raises(PendingDecisionPhaseReportAdapterError, match="digest/date/scope/version"):
        create(data)
    assert not data[3].exists()


def test_adapter_rejects_invalid_phase_paths_and_preserves_existing_report(tmp_path, monkeypatch):
    data = fixture(tmp_path)
    calls = []
    monkeypatch.setattr(
        "rehab_excel.daily_input_pending_decision_phase_adapter.verify_daily_input_projection_preflight_native_readonly",
        lambda *a, **k: calls.append((a, k)),
    )
    with pytest.raises(PendingDecisionPhaseReportAdapterError, match="Invalid operation phase"):
        create(data, operation=data[4])
    outside = tmp_path / "outside.json"
    outside.write_bytes(data[2].read_bytes())
    with pytest.raises(PendingDecisionPhaseReportAdapterError, match="within evidence root"):
        create(data, inventory_path=outside)
    data[3].write_bytes(b"PRESERVE")
    with pytest.raises(PendingDecisionPhaseReportAdapterError, match="already exists"):
        create(data)
    assert calls == []
    assert data[3].read_bytes() == b"PRESERVE"


def test_adapter_source_has_no_workbook_write_or_apply_surface():
    import rehab_excel.daily_input_pending_decision_phase_adapter as module
    source = inspect.getsource(module)
    assert "ThisWorkbook.Save" not in source
    assert "DailyInputJournalApplyEnabled = True" not in source
    assert not hasattr(module, "enable_apply")
