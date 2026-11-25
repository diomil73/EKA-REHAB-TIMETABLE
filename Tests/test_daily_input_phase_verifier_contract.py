from hashlib import sha256

import pytest

from rehab_excel.daily_input_operation import DailyInputOperation, DailyInputOperationPhase as Phase
from rehab_excel.daily_input_phase_evidence import EvidenceKind
from rehab_excel.daily_input_phase_verifier_contract import (
    PHASE_VERIFIER_REPORT_SCHEMA,
    VERIFIER_CONTRACTS,
    DailyInputPhaseVerifierReport,
    PhaseVerifierContractError,
    PhaseVerifierNotImplementedError,
    VerifierImplementationStatus,
    VerificationResult,
    create_phase_verifier_report,
    decode_phase_verifier_report,
    registered_verifier_contracts,
    validate_phase_verifier_report,
)


def operation():
    return DailyInputOperation(
        "op-verifier-001", "a" * 64, "2026-11-25",
        phase=Phase.INPUT_AND_LEDGER_COMMITTED, revision=1,
    )


def prepared_operation():
    return DailyInputOperation("op-verifier-001", "a" * 64, "2026-11-25")


def report_fields(**overrides):
    contract = VERIFIER_CONTRACTS[EvidenceKind.INPUT_AND_LEDGER_REOPEN]
    value = {
        "report_id": "report-001",
        "operation_id": "op-verifier-001",
        "source_workbook_sha256": "a" * 64,
        "service_date": "2026-11-25",
        "journal_revision": 1,
        "journal_phase": Phase.INPUT_AND_LEDGER_COMMITTED.value,
        "evidence_kind": EvidenceKind.INPUT_AND_LEDGER_REOPEN.value,
        "verifier_id": contract.verifier_id,
        "contract_version": contract.contract_version,
        "verification_method": contract.verification_method,
        "verification_result": VerificationResult.VERIFIED.value,
        "evidence_artifact": "op-verifier-001/input-ledger-report.json",
        "evidence_sha256": sha256(b"synthetic report evidence").hexdigest(),
        "inspected_scope": list(contract.required_scope),
        "native_excel_version": "SYNTHETIC-EXCEL-VERSION",
        "package_audit_sha256": None,
        "prepared_snapshot_sha256": None,
        "verified_at_utc": "2026-11-25T11:00:00Z",
    }
    value.update(overrides)
    return value


def test_registry_covers_every_evidence_kind_exactly_once():
    contracts = registered_verifier_contracts()
    assert len(contracts) == len(EvidenceKind)
    assert {item.evidence_kind for item in contracts} == set(EvidenceKind)
    assert len({item.verifier_id for item in contracts}) == len(EvidenceKind)


def test_all_verifiers_remain_explicitly_unimplemented():
    assert all(
        item.implementation_status == VerifierImplementationStatus.NOT_IMPLEMENTED
        for item in registered_verifier_contracts()
    )


def test_registry_is_immutable():
    with pytest.raises(TypeError):
        VERIFIER_CONTRACTS[EvidenceKind.INPUT_AND_LEDGER_REOPEN] = None


def test_report_canonical_round_trip_and_digest_validation():
    report = create_phase_verifier_report(**report_fields())
    assert report.schema == PHASE_VERIFIER_REPORT_SCHEMA
    encoded = report.to_json()
    assert decode_phase_verifier_report(encoded) == report
    assert decode_phase_verifier_report(encoded.encode("utf-8")) == report


def test_report_digest_detects_tampering():
    report = create_phase_verifier_report(**report_fields())
    changed = report.to_dict()
    changed["native_excel_version"] = "SPOOFED-VERSION"
    with pytest.raises(PhaseVerifierContractError, match="SHA-256 mismatch"):
        decode_phase_verifier_report(changed)


def test_report_rejects_missing_extra_and_caller_boolean_fields():
    fields = report_fields()
    fields.pop("verifier_id")
    with pytest.raises(PhaseVerifierContractError, match="missing or unexpected"):
        create_phase_verifier_report(**fields)
    with pytest.raises(PhaseVerifierContractError, match="missing or unexpected"):
        create_phase_verifier_report(**report_fields(caller_verified=True))


@pytest.mark.parametrize("locator", ["../outside.json", "/absolute/report.json", "C:/report.json", "op\\report.json"])
def test_report_rejects_unsafe_artifact_locators(locator):
    with pytest.raises(PhaseVerifierContractError, match="evidence_artifact"):
        create_phase_verifier_report(**report_fields(evidence_artifact=locator))


def test_report_rejects_identity_drift_before_verifier_availability_check():
    report = create_phase_verifier_report(**report_fields(operation_id="different-operation"))
    with pytest.raises(PhaseVerifierContractError, match="operation identity/revision"):
        validate_phase_verifier_report(
            report, operation(),
            previous_operation=prepared_operation(),
        )


def test_report_rejects_unrequired_evidence_kind():
    report = create_phase_verifier_report(**report_fields(
        evidence_kind=EvidenceKind.PROJECTIONS_REOPEN.value,
        verifier_id=VERIFIER_CONTRACTS[EvidenceKind.PROJECTIONS_REOPEN].verifier_id,
        verification_method=VERIFIER_CONTRACTS[EvidenceKind.PROJECTIONS_REOPEN].verification_method,
        inspected_scope=list(VERIFIER_CONTRACTS[EvidenceKind.PROJECTIONS_REOPEN].required_scope),
    ))
    with pytest.raises(PhaseVerifierContractError, match="not required"):
        validate_phase_verifier_report(
            report, operation(),
            required_evidence=(EvidenceKind.INPUT_AND_LEDGER_REOPEN,),
        )


def test_report_rejects_contract_scope_or_method_drift():
    report = create_phase_verifier_report(**report_fields(
        inspected_scope=["sheet:DAILY_INPUT"],
    ))
    with pytest.raises(PhaseVerifierContractError, match="scope differs"):
        validate_phase_verifier_report(
            report, operation(),
            required_evidence=(EvidenceKind.INPUT_AND_LEDGER_REOPEN,),
        )


def test_positive_phase_report_fails_closed_until_native_verifier_is_implemented():
    report = create_phase_verifier_report(**report_fields())
    with pytest.raises(PhaseVerifierNotImplementedError, match="No executable independent verifier"):
        validate_phase_verifier_report(
            report, operation(),
            required_evidence=(EvidenceKind.INPUT_AND_LEDGER_REOPEN,),
        )


def test_rejected_report_cannot_satisfy_positive_phase_obligation():
    report = create_phase_verifier_report(**report_fields(
        verification_result=VerificationResult.REJECTED.value,
    ))
    with pytest.raises(PhaseVerifierContractError, match="Only VERIFIED"):
        validate_phase_verifier_report(
            report, operation(),
            required_evidence=(EvidenceKind.INPUT_AND_LEDGER_REOPEN,),
        )


def test_rollback_contract_requires_native_package_and_prepared_snapshot_digests():
    contract = VERIFIER_CONTRACTS[EvidenceKind.FULL_WORKBOOK_ROLLBACK]
    assert contract.requires_native_excel is True
    assert contract.requires_package_audit is True
    assert contract.requires_prepared_snapshot_binding is True

def test_rollback_report_requires_package_audit_and_prepared_snapshot_hashes():
    contract = VERIFIER_CONTRACTS[EvidenceKind.FULL_WORKBOOK_ROLLBACK]
    rollback_operation = DailyInputOperation(
        "op-verifier-rollback", "a" * 64, "2026-11-25",
        phase=Phase.ROLLED_BACK, revision=2,
        failure_code="SYNTHETIC_FAILURE", rollback_verified=True,
    )
    fields = {
        "report_id": "rollback-report-001",
        "operation_id": rollback_operation.operation_id,
        "source_workbook_sha256": rollback_operation.source_sha256,
        "service_date": rollback_operation.service_date,
        "journal_revision": rollback_operation.revision,
        "journal_phase": rollback_operation.phase.value,
        "evidence_kind": EvidenceKind.FULL_WORKBOOK_ROLLBACK.value,
        "verifier_id": contract.verifier_id,
        "contract_version": contract.contract_version,
        "verification_method": contract.verification_method,
        "verification_result": VerificationResult.VERIFIED.value,
        "evidence_artifact": "op-verifier-rollback/rollback-report.json",
        "evidence_sha256": sha256(b"synthetic rollback evidence").hexdigest(),
        "inspected_scope": list(contract.required_scope),
        "native_excel_version": "SYNTHETIC-EXCEL-VERSION",
        "package_audit_sha256": None,
        "prepared_snapshot_sha256": None,
        "verified_at_utc": "2026-11-25T11:30:00Z",
    }
    report = create_phase_verifier_report(**fields)
    with pytest.raises(PhaseVerifierContractError, match="Package audit digest"):
        validate_phase_verifier_report(
            report, rollback_operation,
            previous_operation=DailyInputOperation(
                rollback_operation.operation_id, rollback_operation.source_sha256,
                rollback_operation.service_date, phase=Phase.FAILED, revision=1,
                failure_code="SYNTHETIC_FAILURE",
            ),
        )
    fields["package_audit_sha256"] = "b" * 64
    report = create_phase_verifier_report(**fields)
    with pytest.raises(PhaseVerifierContractError, match="PREPARED snapshot digest"):
        validate_phase_verifier_report(
            report, rollback_operation,
            required_evidence=(EvidenceKind.FULL_WORKBOOK_ROLLBACK,),
        )

