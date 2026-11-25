"""Synthetic-only, non-authorizing verifier attestation requirements tests."""
import base64
from dataclasses import replace
from hashlib import sha256
import json

import pytest

from rehab_excel.daily_input_phase_evidence import EvidenceKind
from rehab_excel.daily_input_phase_verifier_contract import (
    VERIFIER_CONTRACTS,
    create_phase_verifier_report,
)
from rehab_excel.daily_input_verifier_attestation import (
    ATTESTATION_SCHEMA,
    VerifierAttestationError,
    VerifierAttestationUnavailable,
    decode_untrusted_attestation_claim,
    require_trusted_verifier_attestation,
)


def sample_report():
    contract = VERIFIER_CONTRACTS[EvidenceKind.INPUT_AND_LEDGER_REOPEN]
    return create_phase_verifier_report(
        report_id="fake-report", operation_id="fake-operation",
        source_workbook_sha256="a" * 64, service_date="2026-11-25",
        journal_revision=1, journal_phase="INPUT_AND_LEDGER_COMMITTED",
        evidence_kind=EvidenceKind.INPUT_AND_LEDGER_REOPEN.value,
        verifier_id=contract.verifier_id, contract_version=contract.contract_version,
        verification_method=contract.verification_method,
        verification_result="VERIFIED", evidence_artifact="fake/output.xlsm",
        evidence_sha256="b" * 64, inspected_scope=list(contract.required_scope),
        native_excel_version="SYNTHETIC", package_audit_sha256=None,
        prepared_snapshot_sha256=None, verified_at_utc="2026-11-25T11:00:00Z",
    )


def sample_claim(**changes):
    report = sample_report()
    fields = dict(
        schema=ATTESTATION_SCHEMA,
        verifier_id=report.verifier_id, contract_version=report.contract_version,
        evidence_kind=report.evidence_kind,
        report_sha256=report.report_sha256,
        verifier_build_sha256=sha256(b"fictional-untrusted-build").hexdigest(),
        source_workbook_sha256=report.source_workbook_sha256,
        operation_id=report.operation_id, service_date=report.service_date,
        journal_revision=report.journal_revision, journal_phase=report.journal_phase,
        signer_key_id="fictional-untrusted-key",
        signature_algorithm="Ed25519",
        signature_b64=base64.b64encode(b"\x00" * 64).decode("ascii"),
        issued_at_utc="2026-11-25T10:00:00Z",
        expires_at_utc="2026-11-26T10:00:00Z",
    )
    fields.update(changes)
    return json.dumps(fields, sort_keys=True, separators=(",", ":"))


def test_valid_looking_self_asserted_claim_is_always_denied():
    report = sample_report()
    claim = decode_untrusted_attestation_claim(sample_claim())
    with pytest.raises(VerifierAttestationUnavailable, match="No independently pinned"):
        require_trusted_verifier_attestation(report, claim)


def test_missing_claim_is_denied():
    with pytest.raises(VerifierAttestationUnavailable):
        require_trusted_verifier_attestation(sample_report())


@pytest.mark.parametrize("changes", [
    {"verifier_id": "fake-verifier"},
    {"contract_version": "contract-v9"},
    {"source_workbook_sha256": "c" * 64},
    {"journal_revision": 2},
    {"report_sha256": "d" * 64},
    {"operation_id": "other-operation"},
])
def test_rejects_mismatched_identity(changes):
    claim = decode_untrusted_attestation_claim(sample_claim(**changes))
    with pytest.raises(VerifierAttestationError, match="Claim differs"):
        require_trusted_verifier_attestation(sample_report(), claim)


@pytest.mark.parametrize("changes", [
    {"signature_algorithm": "none"},
    {"signature_b64": "AA=="},
    {"signature_b64": "not-base64"},
    {"journal_revision": True},
    {"expires_at_utc": "2026-11-25T10:00:00Z"},
    {"issued_at_utc": "2026-11-25T10:00:00+00:00"},
])
def test_rejects_malformed_claim(changes):
    with pytest.raises(VerifierAttestationError):
        decode_untrusted_attestation_claim(sample_claim(**changes))


def test_canonical_only_no_extra_fields_or_caller_boolean():
    with pytest.raises(VerifierAttestationError, match="Noncanonical"):
        decode_untrusted_attestation_claim(json.dumps(json.loads(sample_claim()), indent=2))
    with pytest.raises(VerifierAttestationError, match="fields missing or unexpected"):
        decode_untrusted_attestation_claim(sample_claim()[:-1] + ',"verified":true}')
    with pytest.raises(VerifierAttestationError):
        decode_untrusted_attestation_claim({})


def test_forged_direct_dataclass_report_is_rejected():
    report = replace(sample_report(), evidence_sha256="f" * 64)
    claim = decode_untrusted_attestation_claim(sample_claim())
    with pytest.raises(VerifierAttestationError):
        require_trusted_verifier_attestation(report, claim)
