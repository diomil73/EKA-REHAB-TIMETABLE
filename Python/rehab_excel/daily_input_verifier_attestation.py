"""Fail-closed attestation *requirements* for DAILY_INPUT verifier reports.

This module is deliberately not a trust provider, certificate validator, signer,
signature verifier, key store, or a runtime approval mechanism. It validates
the untrusted claimed provenance envelope and records the missing trusted
external prerequisites. Neither a JSON digest nor a Git commit authenticates
the producer of a report. No producer may self-attest.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import re

from .daily_input_phase_evidence import EvidenceKind
from .daily_input_phase_verifier_contract import (
    DailyInputPhaseVerifierReport,
    VERIFIER_CONTRACTS,
)

ATTESTATION_SCHEMA = "diomil-verifier-attestation-requirements-v1"
_HEX = re.compile(r"^[0-9a-f]{64}$")
_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$")
_FIELDS = frozenset({
    "schema", "verifier_id", "contract_version", "evidence_kind",
    "report_sha256", "verifier_build_sha256", "source_workbook_sha256",
    "operation_id", "service_date", "journal_revision", "journal_phase",
    "signer_key_id", "signature_algorithm", "signature_b64",
    "issued_at_utc", "expires_at_utc",
})


class VerifierAttestationError(ValueError):
    """Untrusted attestation envelope is invalid or cannot be authenticated."""


class VerifierAttestationUnavailable(VerifierAttestationError):
    """No independently administered cryptographic trust service is configured."""


def _utc(text: object) -> datetime:
    if type(text) is not str or not text.endswith("Z"):
        raise VerifierAttestationError("Attestation timestamp must be canonical UTC")
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as exc:
        raise VerifierAttestationError("Invalid attestation timestamp") from exc
    if parsed.utcoffset() is None or parsed.utcoffset().total_seconds() != 0:
        raise VerifierAttestationError("Attestation timestamp must be UTC")
    if parsed.isoformat().replace("+00:00", "Z") != text:
        raise VerifierAttestationError("Attestation timestamp must be canonical")
    return parsed


@dataclass(frozen=True)
class UntrustedVerifierAttestationClaim:
    """A report-bound claim, never a verification verdict."""

    schema: str
    verifier_id: str
    contract_version: str
    evidence_kind: str
    report_sha256: str
    verifier_build_sha256: str
    source_workbook_sha256: str
    operation_id: str
    service_date: str
    journal_revision: int
    journal_phase: str
    signer_key_id: str
    signature_algorithm: str
    signature_b64: str
    issued_at_utc: str
    expires_at_utc: str

    def __post_init__(self):
        if self.schema != ATTESTATION_SCHEMA:
            raise VerifierAttestationError("Unknown attestation schema")
        for name in (
            "verifier_id", "contract_version", "operation_id",
            "journal_phase", "signer_key_id",
        ):
            v = getattr(self, name)
            if type(v) is not str or not _ID.fullmatch(v):
                raise VerifierAttestationError("Invalid attestation identifier")
        for name in ("report_sha256", "verifier_build_sha256", "source_workbook_sha256"):
            v = getattr(self, name)
            if type(v) is not str or not _HEX.fullmatch(v):
                raise VerifierAttestationError("Invalid attestation digest")
        if type(self.journal_revision) is not int or self.journal_revision < 0:
            raise VerifierAttestationError("Invalid journal revision")
        if self.signature_algorithm != "Ed25519":
            raise VerifierAttestationError("Unsupported attestation algorithm")
        if type(self.signature_b64) is not str:
            raise VerifierAttestationError("Invalid attestation signature")
        # Signature syntax/length alone is never cryptographic validation.
        import base64
        try:
            raw = base64.b64decode(self.signature_b64, validate=True)
        except (ValueError, base64.binascii.Error) as exc:
            raise VerifierAttestationError("Invalid attestation signature encoding") from exc
        if len(raw) != 64 or base64.b64encode(raw).decode("ascii") != self.signature_b64:
            raise VerifierAttestationError("Invalid attestation signature length/encoding")
        try:
            EvidenceKind(self.evidence_kind)
        except ValueError as exc:
            raise VerifierAttestationError("Unknown evidence kind") from exc
        if _utc(self.issued_at_utc) >= _utc(self.expires_at_utc):
            raise VerifierAttestationError("Attestation validity window invalid")


def decode_untrusted_attestation_claim(value: str | bytes) -> UntrustedVerifierAttestationClaim:
    """Parse only canonical UTF-8 JSON; never trust a decoded object."""
    if type(value) not in (str, bytes):
        raise VerifierAttestationError("Attestation must be raw canonical JSON")
    try:
        raw = value.decode("utf-8") if type(value) is bytes else value
        data = json.loads(raw)
    except (UnicodeError, ValueError) as exc:
        raise VerifierAttestationError("Invalid attestation JSON") from exc
    if type(data) is not dict or set(data) != _FIELDS:
        raise VerifierAttestationError("Attestation fields missing or unexpected")
    canonical = json.dumps(data, sort_keys=True, separators=(",", ":"), ensure_ascii=True, allow_nan=False)
    if raw != canonical:
        raise VerifierAttestationError("Noncanonical attestation JSON")
    try:
        return UntrustedVerifierAttestationClaim(**data)
    except TypeError as exc:
        raise VerifierAttestationError("Invalid attestation fields") from exc


def require_trusted_verifier_attestation(
    report: DailyInputPhaseVerifierReport,
    claim: UntrustedVerifierAttestationClaim | None = None,
) -> None:
    """Always deny until an independent authenticated trust provider is installed.

    This is not a stub whose return value can be mistaken for approval:
    even a syntactically correct, perfectly matching self-asserted claim
    raises. Future work must provide a *separate* trusted signature verifier,
    externally pinned public key/build manifest, revocation/expiry policy,
    private-key custody, and execution identity verification. Do not remove
    this guard merely by passing an input flag.
    """
    if type(report) is not DailyInputPhaseVerifierReport:
        raise VerifierAttestationError("Invalid phase verifier report")
    if claim is None or type(claim) is not UntrustedVerifierAttestationClaim:
        raise VerifierAttestationUnavailable("Trusted verifier attestation unavailable")
    from .daily_input_phase_verifier_contract import (
        PhaseVerifierContractError,
        decode_phase_verifier_report,
    )

    # Recompute report self-digest; a dataclass may otherwise be modified.
    try:
        decode_phase_verifier_report(report.to_json())
    except PhaseVerifierContractError as exc:
        raise VerifierAttestationError("Phase report integrity failed") from exc
    try:
        contract = VERIFIER_CONTRACTS[EvidenceKind(claim.evidence_kind)]
    except (ValueError, KeyError) as exc:
        raise VerifierAttestationError("Unknown evidence contract") from exc
    expected = (
        report.verifier_id, report.contract_version, report.evidence_kind,
        report.report_sha256, report.source_workbook_sha256,
        report.operation_id, report.service_date,
        report.journal_revision, report.journal_phase,
    )
    actual = (
        claim.verifier_id, claim.contract_version, claim.evidence_kind,
        claim.report_sha256, claim.source_workbook_sha256,
        claim.operation_id, claim.service_date,
        claim.journal_revision, claim.journal_phase,
    )
    if actual != expected or (contract.verifier_id, contract.contract_version) != (
        report.verifier_id, report.contract_version
    ):
        raise VerifierAttestationError("Claim differs from pinned report identity")
    raise VerifierAttestationUnavailable("No independently pinned verifier attestation trust root")
