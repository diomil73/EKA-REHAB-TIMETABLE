"""Synthetic tests for non-authorizing PROJECTIONS_REOPEN report adapter."""
from dataclasses import replace
from datetime import date, time
from hashlib import sha256
from pathlib import Path

import pytest

import rehab_excel.daily_input_projections_phase_adapter as adapter
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_phase_verifier_contract import (
    PhaseVerifierNotImplementedError, decode_phase_verifier_report,
    validate_phase_verifier_report,
)
from rehab_excel.daily_input_projection_preflight import DailyInputProjectionPreflight
from rehab_excel.daily_input_projection_preflight_native import (
    DailyInputProjectionPreflightNativeReadback,
)
from rehab_excel.daily_input_projection_native import DailyInputProjectionNativeReadback
from rehab_excel.daily_input_source_native_capture import (
    DailyInputSourceNativeCapture, NativeSourceSheetResult, REQUIRED_SOURCE_SHEETS,
)
from rehab_excel.daily_input_reader import DailyInputReadResult, DailyInputSourceChange
from rehab_excel.daily_input_therapist_projection_expected import TherapistDailyExpectedGrid
from rehab_excel.reader import WorkbookSettings


DAY = date(2026, 11, 23)


def fixture(tmp_path):
    root = tmp_path / "evidence"
    root.mkdir()
    workbook = root / "synthetic-projection.xlsm"
    workbook.write_bytes(b"SYNTHETIC disposable workbook only")
    output = root / "projections-report.json"
    digest = sha256(workbook.read_bytes()).hexdigest()

    initial = DailyInputOperation("projection-demo-001", "a" * 64, DAY.isoformat())
    previous = advance_daily_input_operation(initial, Phase.INPUT_AND_LEDGER_COMMITTED)
    operation = advance_daily_input_operation(previous, Phase.PROJECTIONS_VERIFIED)

    daily = DailyInputReadResult(
        target_date=DAY, absences=(), therapist_rows_used=1, patient_rows_used=0,
        source_changes=(DailyInputSourceChange(
            "change:0001", 7, "therapist_absence", "SYNTH-PT",
            "SYNTH-PT", "synthetic", True,
        ),),
    )
    grid = TherapistDailyExpectedGrid(
        service_date=DAY, source_workbook_sha256=digest,
        provider_headers=("",) * 18, cells=(), expected_assignments=(),
        occupied_cells=0,
    )
    settings = WorkbookSettings(
        therapist_names=("SYNTH-PT",), standard_timeslots=(time(8, 30),),
        reclined_timeslots=(), day_patterns=("ΔΕ",), therapies=("ΦΘ",),
        patient_statuses=("Active",), yes_no_values=("ΝΑΙ", "ΟΧΙ"),
        rooms=("A01",),
    )
    preflight = DailyInputProjectionPreflight(
        service_date=DAY, baseline_date=date(2026, 1, 5),
        source_workbook_sha256=digest, patients=(), settings=settings,
        stable_base_entries=(), effective_base_entries=(), recurring_changes=(),
        patient_exits=(), sessions=(), daily_input=daily, replacements=(),
        states=(), students=(), expected_projection=(), expected_therapist_grid=grid,
        approved_shared_memberships=(),
    )
    capture = DailyInputSourceNativeCapture(
        workbook_sha256=digest, native_excel_version="16.0",
        sheets=tuple(NativeSourceSheetResult(name, 1, 1, 1) for name in REQUIRED_SOURCE_SHEETS),
        identity_mirror_rows_checked=0,
    )
    projections = DailyInputProjectionNativeReadback(
        workbook_sha256=digest, service_date=DAY,
        master_projected_rows=0, therapist_grid_cells_checked=126,
        therapist_occupied_cells=0, expected_therapist_assignments=0,
        input_ledger_source_changes=1, input_ledger_active_rows=1,
        master_excel_version="16.0", therapist_excel_version="16.0",
    )
    native = DailyInputProjectionPreflightNativeReadback(
        preflight=preflight, native_readback=projections, source_capture=capture,
    )
    return root, workbook, output, previous, operation, preflight, native


def create(root, workbook, output, previous, operation, preflight, **overrides):
    args = dict(
        previous_operation=previous, operation=operation, preflight=preflight,
        evidence_root=root, report_path=output, report_id="projection-report-001",
    )
    args.update(overrides)
    return adapter.create_projections_phase_report(workbook, **args)


def test_adapter_emits_exact_bound_report_but_cannot_authorize(tmp_path, monkeypatch):
    root, workbook, output, previous, operation, preflight, native = fixture(tmp_path)
    calls = []

    def fake_verify(path, **kwargs):
        calls.append((path, kwargs))
        return native

    monkeypatch.setattr(adapter, "verify_daily_input_projection_preflight_native_readonly", fake_verify)
    report = create(root, workbook, output, previous, operation, preflight)
    assert calls[0][0] == workbook.resolve()
    assert calls[0][1]["preflight"] is preflight
    assert report.operation_id == operation.operation_id
    assert report.journal_phase == operation.phase.value
    assert report.journal_revision == operation.revision
    assert report.source_workbook_sha256 == operation.source_sha256
    assert report.source_workbook_sha256 != sha256(workbook.read_bytes()).hexdigest()
    assert report.evidence_sha256 == sha256(workbook.read_bytes()).hexdigest()
    assert report.evidence_artifact == workbook.name
    assert report.native_excel_version == "16.0"
    assert report.inspected_scope == (
        "sheet:MASTER_SCHEDULE", "affected_therapist_projections",
        "source_input_ledger_linkage",
    )
    assert decode_phase_verifier_report(output.read_text(encoding="utf-8").strip()) == report
    with pytest.raises(PhaseVerifierNotImplementedError, match="trusted verifier attestation"):
        validate_phase_verifier_report(report, operation, previous_operation=previous)


@pytest.mark.parametrize("problem", ["wrong-hash", "wrong-version", "wrong-scope", "mutation"])
def test_adapter_never_publishes_inconsistent_native_evidence(tmp_path, monkeypatch, problem):
    root, workbook, output, previous, operation, preflight, native = fixture(tmp_path)
    if problem == "wrong-hash":
        capture = native.source_capture
        native = DailyInputProjectionPreflightNativeReadback(
            preflight=preflight, native_readback=native.native_readback,
            source_capture=DailyInputSourceNativeCapture(
                workbook_sha256="f" * 64, native_excel_version="16.0",
                sheets=capture.sheets, identity_mirror_rows_checked=0,
            ),
        )
    elif problem == "wrong-version":
        capture = native.source_capture
        native = DailyInputProjectionPreflightNativeReadback(
            preflight=preflight, native_readback=native.native_readback,
            source_capture=DailyInputSourceNativeCapture(
                workbook_sha256=capture.workbook_sha256, native_excel_version="15.0",
                sheets=capture.sheets, identity_mirror_rows_checked=0,
            ),
        )
    elif problem == "wrong-scope":
        p = native.native_readback
        bad_projection = DailyInputProjectionNativeReadback(
            workbook_sha256=p.workbook_sha256, service_date=p.service_date,
            master_projected_rows=1, therapist_grid_cells_checked=p.therapist_grid_cells_checked,
            therapist_occupied_cells=p.therapist_occupied_cells,
            expected_therapist_assignments=p.expected_therapist_assignments,
            input_ledger_source_changes=p.input_ledger_source_changes,
            input_ledger_active_rows=p.input_ledger_active_rows,
            master_excel_version=p.master_excel_version,
            therapist_excel_version=p.therapist_excel_version,
        )
        native = DailyInputProjectionPreflightNativeReadback(
            preflight=preflight, native_readback=bad_projection,
            source_capture=native.source_capture,
        )

    def fake_verify(path, **kwargs):
        if problem == "mutation":
            path.write_bytes(path.read_bytes() + b"changed")
        return native

    monkeypatch.setattr(adapter, "verify_daily_input_projection_preflight_native_readonly", fake_verify)
    with pytest.raises(adapter.ProjectionsPhaseReportAdapterError):
        create(root, workbook, output, previous, operation, preflight)
    assert not output.exists()


def test_adapter_rejects_wrong_phase_or_preflight_identity_before_native(tmp_path, monkeypatch):
    root, workbook, output, previous, operation, preflight, _native = fixture(tmp_path)
    calls = []
    monkeypatch.setattr(
        adapter, "verify_daily_input_projection_preflight_native_readonly",
        lambda *a, **k: calls.append((a, k)),
    )
    with pytest.raises(adapter.ProjectionsPhaseReportAdapterError, match="Invalid phase transition"):
        create(root, workbook, output, previous, previous, preflight)
    assert calls == []
    bad_preflight = replace(preflight, service_date=date(2026, 11, 24))
    with pytest.raises(adapter.ProjectionsPhaseReportAdapterError, match="Preflight/operation"):
        create(root, workbook, output, previous, operation, bad_preflight)
    assert calls == []


def test_adapter_rejects_outside_paths_and_existing_output_before_native(tmp_path, monkeypatch):
    root, workbook, output, previous, operation, preflight, _native = fixture(tmp_path)
    calls = []
    monkeypatch.setattr(
        adapter, "verify_daily_input_projection_preflight_native_readonly",
        lambda *a, **k: calls.append((a, k)),
    )
    outside = tmp_path / "outside.xlsm"
    outside.write_bytes(workbook.read_bytes())
    with pytest.raises(adapter.ProjectionsPhaseReportAdapterError, match="within evidence root"):
        create(root, outside, output, previous, operation, preflight)
    output.write_bytes(b"PRESERVE")
    with pytest.raises(adapter.ProjectionsPhaseReportAdapterError, match="already exists"):
        create(root, workbook, output, previous, operation, preflight)
    assert calls == []
    assert output.read_bytes() == b"PRESERVE"


def test_adapter_does_not_write_workbook_or_enable_apply():
    import inspect
    source = inspect.getsource(adapter)
    assert "DailyInputJournalApplyEnabled = True" not in source
    assert "ThisWorkbook.Save" not in source
    assert not hasattr(adapter, "enable_apply")
