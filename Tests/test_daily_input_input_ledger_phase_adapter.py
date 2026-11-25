"""Non-authorizing INPUT_AND_LEDGER phase report adapter tests, synthetic only."""
from dataclasses import replace
from datetime import date

import pytest

from rehab_excel.authoritative_commit import file_sha256
from rehab_excel.daily_input_input_ledger_native import (
    InputLedgerNativeReadback, InputLedgerNativeReadbackError,
)
from rehab_excel.daily_input_input_ledger_phase_adapter import (
    InputLedgerPhaseReportAdapterError, create_input_ledger_phase_report,
)
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_phase_evidence import EvidenceKind
from rehab_excel.daily_input_phase_verifier_contract import (
    PhaseVerifierNotImplementedError, decode_phase_verifier_report,
    validate_phase_verifier_report,
)
from rehab_excel.daily_input_verifier_attestation import (
    VerifierAttestationUnavailable, require_trusted_verifier_attestation,
)


def fixture(tmp_path):
    root = tmp_path / "evidence"
    root.mkdir()
    workbook = root / "fictional-snapshot.xlsm"
    workbook.write_bytes(b"SYNTHETIC and NOT a real XLSM")
    output = root / "input-ledger-report.json"
    previous = DailyInputOperation("input-ledger-demo-001", "a" * 64, "2026-11-25")
    operation = advance_daily_input_operation(previous, Phase.INPUT_AND_LEDGER_COMMITTED)
    return root, workbook, output, previous, operation


def install_diagnostic(monkeypatch, workbook, *, mutate=False, raise_error=False, wrong_hash=False):
    import rehab_excel.daily_input_input_ledger_phase_adapter as adapter
    calls = []
    def fake_verify(path, **kwargs):
        calls.append((path, kwargs))
        if raise_error:
            raise InputLedgerNativeReadbackError("fake failure")
        sha = file_sha256(path)
        if mutate:
            path.write_bytes(b"CHANGED AFTER INSPECTION")
        return InputLedgerNativeReadback(
            service_date=kwargs["service_date"],
            workbook_sha256="f" * 64 if wrong_hash else sha,
            source_change_count=1, active_ledger_count=1,
            native_excel_version="SYNTHETIC-EXCEL-16",
        )
    monkeypatch.setattr(adapter, "verify_input_ledger_native_readonly", fake_verify)
    return calls


def create(root, workbook, output, previous, operation, **overrides):
    args = dict(
        previous_operation=previous, operation=operation,
        validated_changes=(object(),), patients=(), sessions=(),
        evidence_root=root, report_path=output, report_id="input-ledger-report-001",
    )
    args.update(overrides)
    return create_input_ledger_phase_report(workbook, **args)


def test_adapter_emits_exact_bound_canonical_report_but_cannot_authorize(tmp_path, monkeypatch):
    root, workbook, output, previous, operation = fixture(tmp_path)
    calls = install_diagnostic(monkeypatch, workbook)
    original = workbook.read_bytes()
    report = create(root, workbook, output, previous, operation)
    assert len(calls) == 1
    assert calls[0][0] == workbook.resolve()
    assert calls[0][1]["validated_changes"] == (calls[0][1]["validated_changes"][0],)
    assert report.operation_id == operation.operation_id
    assert report.journal_phase == operation.phase.value
    assert report.journal_revision == operation.revision
    assert report.source_workbook_sha256 == operation.source_sha256
    assert report.evidence_kind == EvidenceKind.INPUT_AND_LEDGER_REOPEN.value
    assert report.evidence_sha256 == file_sha256(workbook)
    assert report.evidence_artifact == workbook.name
    assert report.native_excel_version == "SYNTHETIC-EXCEL-16"
    assert decode_phase_verifier_report(output.read_text(encoding="utf-8").strip()) == report
    assert workbook.read_bytes() == original
    with pytest.raises(PhaseVerifierNotImplementedError, match="trusted verifier attestation"):
        validate_phase_verifier_report(report, operation, previous_operation=previous)
    with pytest.raises(VerifierAttestationUnavailable):
        require_trusted_verifier_attestation(report)


@pytest.mark.parametrize("problem", ["mutation", "diagnostic", "bad-sha"])
def test_adapter_never_publishes_after_bad_diagnostic(tmp_path, monkeypatch, problem):
    root, workbook, output, previous, operation = fixture(tmp_path)
    install_diagnostic(
        monkeypatch, workbook, mutate=problem == "mutation",
        raise_error=problem == "diagnostic", wrong_hash=problem == "bad-sha",
    )
    with pytest.raises(InputLedgerPhaseReportAdapterError):
        create(root, workbook, output, previous, operation)
    assert not output.exists()


def test_adapter_rejects_existing_report_and_never_overwrites(tmp_path, monkeypatch):
    root, workbook, output, previous, operation = fixture(tmp_path)
    calls = install_diagnostic(monkeypatch, workbook)
    output.write_bytes(b"PRESERVE EXISTING")
    with pytest.raises(InputLedgerPhaseReportAdapterError, match="already exists"):
        create(root, workbook, output, previous, operation)
    assert output.read_bytes() == b"PRESERVE EXISTING"
    assert calls == []


def test_adapter_rejects_outside_workbook_before_running_verifier(tmp_path, monkeypatch):
    root, workbook, output, previous, operation = fixture(tmp_path)
    external = tmp_path / "outside.xlsm"
    external.write_bytes(workbook.read_bytes())
    calls = install_diagnostic(monkeypatch, workbook)
    with pytest.raises(InputLedgerPhaseReportAdapterError, match="within evidence root"):
        create(root, external, output, previous, operation)
    assert calls == []


def test_adapter_rejects_illegal_phase_transition_before_verifying(tmp_path, monkeypatch):
    root, workbook, output, previous, operation = fixture(tmp_path)
    calls = install_diagnostic(monkeypatch, workbook)
    with pytest.raises(InputLedgerPhaseReportAdapterError, match="Invalid phase transition"):
        create(root, workbook, output, previous, replace(operation, revision=2))
    assert calls == []


def test_adapter_rejects_zero_changes_and_wrong_service_date(tmp_path, monkeypatch):
    root, workbook, output, previous, operation = fixture(tmp_path)
    calls = install_diagnostic(monkeypatch, workbook)
    with pytest.raises(InputLedgerPhaseReportAdapterError, match="preflight"):
        create(root, workbook, output, previous, operation, validated_changes=())
    assert calls == []
    with pytest.raises(InputLedgerPhaseReportAdapterError, match="service date"):
        create(root, workbook, output, previous, replace(operation, service_date="invalid"))
    assert calls == []


def test_adapter_does_not_write_workbook_or_enable_apply():
    import inspect
    import rehab_excel.daily_input_input_ledger_phase_adapter as adapter
    source = inspect.getsource(adapter)
    assert "DailyInputJournalApplyEnabled = True" not in source
    assert "ThisWorkbook.Save" not in source
    assert not hasattr(adapter, "enable_apply")
