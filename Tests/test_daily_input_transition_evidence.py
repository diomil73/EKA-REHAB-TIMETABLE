from dataclasses import replace
from hashlib import sha256

import pytest

from rehab_excel.daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_phase_evidence import EvidenceKind, required_evidence_for_transition
from rehab_excel.daily_input_phase_verifier_contract import (
    VERIFIER_CONTRACTS,
    PhaseVerifierNotImplementedError,
    VerifierImplementationStatus,
    VerificationResult,
    create_phase_verifier_report,
)
from rehab_excel.daily_input_transaction_receipt import create_transaction_receipt
from rehab_excel import daily_input_phase_verifier_contract as verifier_contract_module
from rehab_excel.daily_input_transition_evidence import (
    TransitionEvidenceBundleError,
    validate_transition_evidence_bundle,
)


def _pending_transition():
    previous = DailyInputOperation("op-bundle-001", "a" * 64, "2026-11-25")
    current = advance_daily_input_operation(previous, Phase.INPUT_AND_LEDGER_COMMITTED)
    current = advance_daily_input_operation(current, Phase.PROJECTIONS_VERIFIED)
    current = advance_daily_input_operation(current, Phase.PENDING_OPERATOR_DECISION)
    candidate = advance_daily_input_operation(
        current, Phase.REPLACEMENTS_RESOLVED, operator_resolution="manual_completed"
    )
    return current, candidate


def _build_bundle(tmp_path):
    previous, operation = _pending_transition()
    required = required_evidence_for_transition(previous, operation)
    reports = []
    receipts = []
    registry = {}
    predecessor = "b" * 64
    for index, kind in enumerate(required, start=1):
        contract = VERIFIER_CONTRACTS[kind]
        payload = f"synthetic evidence {kind.value}".encode()
        artifact = f"{operation.operation_id}/{index}.bin"
        target = tmp_path / artifact
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(payload)
        digest = sha256(payload).hexdigest()
        reports.append(create_phase_verifier_report(
            report_id=f"report-{index:03d}",
            operation_id=operation.operation_id,
            source_workbook_sha256=operation.source_sha256,
            service_date=operation.service_date,
            journal_revision=operation.revision,
            journal_phase=operation.phase.value,
            evidence_kind=kind.value,
            verifier_id=contract.verifier_id,
            contract_version=contract.contract_version,
            verification_method=contract.verification_method,
            verification_result=VerificationResult.VERIFIED.value,
            evidence_artifact=artifact,
            evidence_sha256=digest,
            inspected_scope=list(contract.required_scope),
            native_excel_version="SYNTHETIC-EXCEL" if contract.requires_native_excel else None,
            package_audit_sha256=None,
            prepared_snapshot_sha256=None,
            verified_at_utc=f"2026-11-25T12:0{index}:00Z",
        ))
        receipt = create_transaction_receipt(
            schema="diomil-transaction-receipt-v1",
            receipt_id=f"receipt-{index:03d}",
            operation_id=operation.operation_id,
            source_workbook_sha256=operation.source_sha256,
            service_date=operation.service_date,
            journal_revision=operation.revision,
            journal_phase=operation.phase.value,
            evidence_kind=kind.value,
            evidence_artifact=artifact,
            evidence_sha256=digest,
            verifier_id=contract.verifier_id,
            verifier_version=contract.contract_version,
            verification_method=contract.verification_method,
            verification_result=VerificationResult.VERIFIED.value,
            verified_at_utc=f"2026-11-25T12:1{index}:00Z",
            previous_receipt_sha256=predecessor,
        )
        receipts.append(receipt)
        predecessor = receipt.receipt_sha256
        registry[contract.verifier_id] = contract.contract_version
    return previous, operation, tuple(reports), tuple(receipts), registry


def test_transition_policy_requires_both_operator_resolution_and_reopen():
    previous, operation = _pending_transition()
    assert required_evidence_for_transition(previous, operation) == (
        EvidenceKind.OPERATOR_RESOLUTION,
        EvidenceKind.REPLACEMENTS_REOPEN,
    )


def test_bundle_rejects_missing_report_before_any_verifier_can_be_authorized(tmp_path):
    previous, operation, reports, receipts, registry = _build_bundle(tmp_path)
    with pytest.raises(TransitionEvidenceBundleError, match="Report evidence set"):
        validate_transition_evidence_bundle(
            previous, operation, reports=reports[:1], receipts=receipts,
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry=registry, evidence_root=tmp_path,
        )


def test_bundle_rejects_duplicate_receipt_kind(tmp_path):
    previous, operation, reports, receipts, registry = _build_bundle(tmp_path)
    with pytest.raises(TransitionEvidenceBundleError, match="Duplicate transaction receipt"):
        validate_transition_evidence_bundle(
            previous, operation, reports=reports, receipts=(receipts[0], receipts[0]),
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry=registry, evidence_root=tmp_path,
        )


def test_bundle_rejects_caller_selected_subset_and_extra_evidence(tmp_path):
    previous, operation, reports, receipts, registry = _build_bundle(tmp_path)
    with pytest.raises(TypeError):
        validate_transition_evidence_bundle(
            previous, operation, required_evidence=(EvidenceKind.OPERATOR_RESOLUTION,),
            reports=reports, receipts=receipts,
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry=registry, evidence_root=tmp_path,
        )
    extra_report = replace(reports[0], evidence_kind=EvidenceKind.INPUT_AND_LEDGER_REOPEN.value)
    with pytest.raises(TransitionEvidenceBundleError, match="Duplicate phase verifier report"):
        validate_transition_evidence_bundle(
            previous, operation, reports=(*reports, extra_report), receipts=receipts,
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry=registry, evidence_root=tmp_path,
        )


def test_bundle_fails_closed_while_independent_verifiers_are_unimplemented(tmp_path):
    previous, operation, reports, receipts, registry = _build_bundle(tmp_path)
    with pytest.raises(PhaseVerifierNotImplementedError, match="No executable independent verifier"):
        validate_transition_evidence_bundle(
            previous, operation, reports=reports, receipts=receipts,
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry=registry, evidence_root=tmp_path,
        )


def test_bundle_checks_full_receipt_chain_artifact_bytes_and_exact_order(tmp_path, monkeypatch):
    previous, operation, reports, receipts, registry = _build_bundle(tmp_path)
    contracts = dict(verifier_contract_module.VERIFIER_CONTRACTS)
    for kind, contract in contracts.items():
        if kind in {EvidenceKind.OPERATOR_RESOLUTION, EvidenceKind.REPLACEMENTS_REOPEN}:
            contracts[kind] = replace(
                contract, implementation_status=VerifierImplementationStatus.IMPLEMENTED
            )
    monkeypatch.setattr(verifier_contract_module, "VERIFIER_CONTRACTS", contracts)
    bundle = validate_transition_evidence_bundle(
        previous, operation, reports=reports, receipts=receipts,
        expected_previous_receipt_sha256="b" * 64,
        verifier_registry=registry, evidence_root=tmp_path,
    )
    assert bundle.required_evidence == (
        EvidenceKind.OPERATOR_RESOLUTION, EvidenceKind.REPLACEMENTS_REOPEN
    )
    assert tuple(item.evidence_kind for item in bundle.receipts) == tuple(
        item.value for item in bundle.required_evidence
    )
    # Corrupting the second receipt's artifact must invalidate the whole bundle.
    artifact = tmp_path / receipts[1].evidence_artifact
    artifact.write_bytes(b"tampered")
    with pytest.raises(TransitionEvidenceBundleError, match="Evidence validation failed"):
        validate_transition_evidence_bundle(
            previous, operation, reports=reports, receipts=receipts,
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry=registry, evidence_root=tmp_path,
        )
