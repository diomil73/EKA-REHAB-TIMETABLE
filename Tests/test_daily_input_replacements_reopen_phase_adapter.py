from dataclasses import replace
from datetime import date, time
from hashlib import sha256
import inspect

import pytest

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import (
    DailySessionCancellation, ReplacementAssignment, ReplacementProviderKind,
    Session, SessionCancellationKind,
)
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_pending_decision_inventory import (
    build_pending_decision_inventory, write_pending_decision_inventory,
)
from rehab_excel.daily_input_phase_verifier_contract import (
    PhaseVerifierNotImplementedError, validate_phase_verifier_report,
)
from rehab_excel.daily_input_projection_preflight import DailyInputProjectionPreflight
from rehab_excel.daily_input_replacements_reopen_phase_adapter import (
    ReplacementsReopenPhaseReportAdapterError,
    create_replacements_reopen_phase_report,
)
from rehab_excel.daily_input_source_native_capture import (
    DailyInputSourceNativeCapture, NativeSourceSheetResult, REQUIRED_SOURCE_SHEETS,
)
from rehab_excel.daily_input_therapist_projection_expected import TherapistDailyExpectedGrid
from rehab_excel.daily_input_reader import DailyInputReadResult, DailyInputSourceChange
from rehab_excel.reader import WorkbookSettings


DAY = date(2026, 11, 27)


def _session(sid, patient, therapist, at):
    return Session(sid, patient, therapist, DAY, at, "ΦΘ", False)


def fixture(tmp_path, outcomes=("replacement", "suspension")):
    root = tmp_path / "evidence"
    root.mkdir()
    workbook = root / "final-synthetic.xlsm"
    workbook.write_bytes(b"SYNTHETIC final workbook only")
    inventory_path = root / "pending-inventory.json"
    evidence_path = root / "replacements-evidence.json"
    report_path = root / "replacements-report.json"
    current_hash = sha256(workbook.read_bytes()).hexdigest()

    initial = DailyInputOperation("replacement-reopen-demo-001", "a" * 64, DAY.isoformat())
    input_committed = advance_daily_input_operation(initial, Phase.INPUT_AND_LEDGER_COMMITTED)
    projections_verified = advance_daily_input_operation(input_committed, Phase.PROJECTIONS_VERIFIED)
    pending_operation = advance_daily_input_operation(projections_verified, Phase.PENDING_OPERATOR_DECISION)
    operation = advance_daily_input_operation(
        pending_operation, Phase.REPLACEMENTS_RESOLVED, operator_resolution="manual_completed",
    )

    sessions = (
        _session("S-REPLACE", "P-1", "T-ORIGINAL-1", time(9, 0)),
        _session("S-SUSPEND", "P-2", "T-ORIGINAL-2", time(10, 0)),
    )
    pending_states = tuple(
        DailySessionState(
            item.session_id, item.patient_id, DAY, DailySessionStatus.THERAPIST_ABSENT,
            item.therapist_id, item.start_time, None, None, reason="synthetic absence",
        )
        for item in sessions
    )
    pending_daily = DailyInputReadResult(
        target_date=DAY, absences=(), therapist_rows_used=1, patient_rows_used=0,
        source_changes=(DailyInputSourceChange(
            "change-001", 7, "therapist_absence", "T-ORIGINAL-1",
            "T-ORIGINAL-1", "synthetic absence", True,
        ),),
    )
    settings = WorkbookSettings(
        therapist_names=("T-ORIGINAL-1", "T-ORIGINAL-2", "T-REPLACEMENT"),
        standard_timeslots=(time(9, 0), time(10, 0)), reclined_timeslots=(),
        day_patterns=("ΔΕ",), therapies=("ΦΘ",), patient_statuses=("Active",),
        yes_no_values=("ΝΑΙ", "ΟΧΙ"), rooms=("A01",),
    )
    pending_grid = TherapistDailyExpectedGrid(
        service_date=DAY, source_workbook_sha256="b" * 64,
        provider_headers=("",) * 18, cells=(), expected_assignments=(), occupied_cells=0,
    )
    pending_preflight = DailyInputProjectionPreflight(
        service_date=DAY, baseline_date=date(2026, 1, 5),
        source_workbook_sha256="b" * 64, patients=(), settings=settings,
        stable_base_entries=(), effective_base_entries=(), recurring_changes=(),
        patient_exits=(), sessions=sessions, daily_input=pending_daily,
        replacements=(), states=pending_states, students=(), expected_projection=(),
        expected_therapist_grid=pending_grid, approved_shared_memberships=(),
    )
    inventory = build_pending_decision_inventory(
        projections_verified, pending_operation, pending_preflight,
    )
    write_pending_decision_inventory(
        inventory, evidence_root=root, artifact_path=inventory_path,
    )

    replacements = ()
    cancellations = ()
    final_states = []
    for session, outcome in zip(sessions, outcomes):
        if outcome == "replacement":
            assignment = ReplacementAssignment(
                replacement_id=f"R-{session.session_id}",
                target_session_id=session.session_id,
                patient_id=session.patient_id,
                original_therapist_id=session.therapist_id,
                replacement_therapist_id="T-REPLACEMENT",
                replacement_date=DAY, replacement_time=session.start_time,
                reason="synthetic replacement",
                replacement_provider_kind=ReplacementProviderKind.THERAPIST,
            )
            replacements += (assignment,)
            final_states.append(DailySessionState(
                session.session_id, session.patient_id, DAY, DailySessionStatus.REPLACED,
                session.therapist_id, session.start_time, assignment.replacement_therapist_id,
                assignment.replacement_time, replacement_id=assignment.replacement_id,
                reason=assignment.reason,
                effective_provider_kind=ReplacementProviderKind.THERAPIST,
            ))
        elif outcome == "suspension":
            cancellation = DailySessionCancellation(
                cancellation_id=f"daily-input:{session.session_id}",
                target_session_id=session.session_id, cancellation_date=DAY,
                kind=SessionCancellationKind.TREATMENT_SUSPENDED,
                reason="Αναστολή θεραπείας",
            )
            cancellations += (cancellation,)
            final_states.append(DailySessionState(
                session.session_id, session.patient_id, DAY, DailySessionStatus.CANCELLED,
                session.therapist_id, session.start_time, None, None,
                reason=cancellation.reason, cancellation_id=cancellation.cancellation_id,
                cancellation_kind=SessionCancellationKind.TREATMENT_SUSPENDED,
            ))
        else:
            final_states.append(DailySessionState(
                session.session_id, session.patient_id, DAY, DailySessionStatus.THERAPIST_ABSENT,
                session.therapist_id, session.start_time, None, None,
            ))

    final_daily = DailyInputReadResult(
        target_date=DAY, absences=(), therapist_rows_used=1, patient_rows_used=len(cancellations),
        cancellations=cancellations,
        source_changes=(DailyInputSourceChange(
            "change-final", 7, "therapist_absence", "T-ORIGINAL-1",
            "T-ORIGINAL-1", "synthetic", True,
        ),),
    )
    final_grid = TherapistDailyExpectedGrid(
        service_date=DAY, source_workbook_sha256=current_hash,
        provider_headers=("",) * 18, cells=(), expected_assignments=(), occupied_cells=0,
    )
    preflight = DailyInputProjectionPreflight(
        service_date=DAY, baseline_date=date(2026, 1, 5),
        source_workbook_sha256=current_hash, patients=(), settings=settings,
        stable_base_entries=(), effective_base_entries=(), recurring_changes=(),
        patient_exits=(), sessions=sessions, daily_input=final_daily,
        replacements=replacements, states=tuple(final_states), students=(),
        expected_projection=(), expected_therapist_grid=final_grid,
        approved_shared_memberships=(),
    )
    sheet_names = tuple(REQUIRED_SOURCE_SHEETS) + (("DAILY_REPLACEMENT_FACTS",) if replacements else ())
    capture = DailyInputSourceNativeCapture(
        workbook_sha256=current_hash, native_excel_version="16.0",
        sheets=tuple(NativeSourceSheetResult(name, 1, 1, 1) for name in sheet_names),
        identity_mirror_rows_checked=0,
    )
    return {
        "root": root, "workbook": workbook, "inventory": inventory_path,
        "evidence": evidence_path, "report": report_path,
        "previous": pending_operation, "operation": operation,
        "preflight": preflight, "capture": capture,
    }


def create(data, **overrides):
    args = dict(
        previous_operation=data["previous"], operation=data["operation"],
        preflight=data["preflight"], evidence_root=data["root"],
        evidence_path=data["evidence"], report_path=data["report"],
        report_id="replacements-report-001",
    )
    args.update(overrides)
    return create_replacements_reopen_phase_report(
        data["workbook"], data["inventory"], **args,
    )


def _patch_capture(monkeypatch, capture):
    monkeypatch.setattr(
        "rehab_excel.daily_input_replacements_reopen_phase_adapter.verify_daily_input_source_native_capture_readonly",
        lambda *a, **k: capture,
    )


def test_adapter_publishes_exact_replacement_and_suspension_partition_but_cannot_authorize(tmp_path, monkeypatch):
    data = fixture(tmp_path)
    _patch_capture(monkeypatch, data["capture"])
    report = create(data)
    assert report.operation_id == data["operation"].operation_id
    assert report.journal_phase == Phase.REPLACEMENTS_RESOLVED.value
    assert report.journal_revision == data["operation"].revision
    assert report.source_workbook_sha256 == data["operation"].source_sha256
    assert report.evidence_artifact == data["evidence"].name
    assert report.native_excel_version == "16.0"
    assert report.inspected_scope == (
        "replacement_facts", "suspensions", "operation_identity", "saved_workbook_sha256",
    )
    with pytest.raises(PhaseVerifierNotImplementedError, match="trusted verifier attestation"):
        validate_phase_verifier_report(
            report, data["operation"], previous_operation=data["previous"],
        )


def test_adapter_accepts_all_pending_targets_as_replacements_or_suspensions(tmp_path, monkeypatch):
    data = fixture(tmp_path, outcomes=("replacement", "replacement"))
    _patch_capture(monkeypatch, data["capture"])
    report = create(data)
    assert report.verification_result == "VERIFIED"


def test_adapter_rejects_unresolved_pending_target_before_native_capture(tmp_path, monkeypatch):
    data = fixture(tmp_path, outcomes=("replacement", "unresolved"))
    calls = []
    monkeypatch.setattr(
        "rehab_excel.daily_input_replacements_reopen_phase_adapter.verify_daily_input_source_native_capture_readonly",
        lambda *a, **k: calls.append((a, k)),
    )
    with pytest.raises(ReplacementsReopenPhaseReportAdapterError, match="facts rejected"):
        create(data)
    assert calls == []
    assert not data["evidence"].exists()
    assert not data["report"].exists()


def test_adapter_rejects_pending_target_that_is_both_replaced_and_suspended(tmp_path, monkeypatch):
    data = fixture(tmp_path)
    session = data["preflight"].sessions[0]
    extra = DailySessionCancellation(
        cancellation_id="suspend-also", target_session_id=session.session_id,
        cancellation_date=DAY, kind=SessionCancellationKind.TREATMENT_SUSPENDED,
    )
    daily = replace(
        data["preflight"].daily_input,
        cancellations=(extra, *data["preflight"].daily_input.cancellations),
    )
    data["preflight"] = replace(data["preflight"], daily_input=daily)
    _patch_capture(monkeypatch, data["capture"])
    with pytest.raises(ReplacementsReopenPhaseReportAdapterError, match="facts rejected"):
        create(data)
    assert not data["evidence"].exists()


def test_adapter_rejects_wrong_native_hash_version_or_missing_replacement_sheet(tmp_path, monkeypatch):
    for problem in ("hash", "version", "sheet"):
        case = tmp_path / problem
        case.mkdir()
        data = fixture(case)
        capture = data["capture"]
        if problem == "hash":
            capture = replace(capture, workbook_sha256="f" * 64)
        elif problem == "version":
            capture = replace(capture, native_excel_version="")
        else:
            capture = replace(capture, sheets=tuple(
                item for item in capture.sheets if item.sheet_name != "DAILY_REPLACEMENT_FACTS"
            ))
        _patch_capture(monkeypatch, capture)
        with pytest.raises(ReplacementsReopenPhaseReportAdapterError, match="digest/version|omitted"):
            create(data)
        assert not data["evidence"].exists()
        assert not data["report"].exists()


def test_adapter_rejects_workbook_mutation_during_native_readback(tmp_path, monkeypatch):
    data = fixture(tmp_path)

    def mutate(path, **kwargs):
        path.write_bytes(path.read_bytes() + b"changed")
        return data["capture"]

    monkeypatch.setattr(
        "rehab_excel.daily_input_replacements_reopen_phase_adapter.verify_daily_input_source_native_capture_readonly",
        mutate,
    )
    with pytest.raises(ReplacementsReopenPhaseReportAdapterError, match="digest/version"):
        create(data)
    assert not data["evidence"].exists()


def test_adapter_rejects_inventory_tamper_and_output_overlap(tmp_path, monkeypatch):
    data = fixture(tmp_path)
    data["inventory"].write_bytes(data["inventory"].read_bytes().replace(b"S-REPLACE", b"S-ALTERED"))
    calls = []
    monkeypatch.setattr(
        "rehab_excel.daily_input_replacements_reopen_phase_adapter.verify_daily_input_source_native_capture_readonly",
        lambda *a, **k: calls.append((a, k)),
    )
    with pytest.raises(ReplacementsReopenPhaseReportAdapterError, match="Pending inventory"):
        create(data)
    assert calls == []
    assert not data["evidence"].exists()


def test_adapter_source_has_no_workbook_write_or_apply_surface():
    import rehab_excel.daily_input_replacements_reopen_phase_adapter as module
    source = inspect.getsource(module)
    assert "ThisWorkbook.Save" not in source
    assert "DailyInputJournalApplyEnabled = True" not in source
    assert not hasattr(module, "enable_apply")
