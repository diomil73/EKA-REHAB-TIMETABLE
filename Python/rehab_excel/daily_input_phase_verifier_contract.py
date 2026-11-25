"""Strict DAILY_INPUT phase-verifier contracts and report envelope.

This module is NOT a native Excel verifier. It defines the pinned evidence
contract for each EvidenceKind and validates the shape/binding of a report.
All implementations remain NOT_IMPLEMENTED until an independent adapter
inspects exact saved workbook bytes through native Excel (where required).
It cannot authenticate a report, authorize a journal append, or enable Apply.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from enum import StrEnum
from hashlib import sha256
import json
from pathlib import PurePosixPath
import re
from types import MappingProxyType
from typing import Mapping

from .daily_input_operation import DailyInputOperation
from .daily_input_phase_evidence import (
    EvidenceKind,
    PhaseEvidenceContractError as EvidencePolicyError,
    required_evidence_for_transition,
)

PHASE_VERIFIER_REPORT_SCHEMA = "diomil-phase-verifier-report-v1"
_HEX = re.compile(r"^[0-9a-f]{64}$")
_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$")
_REPORT_FIELDS = frozenset({
    "schema", "report_id", "operation_id", "source_workbook_sha256",
    "service_date", "journal_revision", "journal_phase", "evidence_kind",
    "verifier_id", "contract_version", "verification_method",
    "verification_result", "evidence_artifact", "evidence_sha256",
    "inspected_scope", "native_excel_version", "package_audit_sha256",
    "prepared_snapshot_sha256", "verified_at_utc", "report_sha256",
})


class PhaseVerifierContractError(ValueError):
    """A phase-verifier contract or report is invalid."""


class PhaseVerifierNotImplementedError(PhaseVerifierContractError):
    """No executable independent verifier is registered for this evidence kind."""


class VerifierImplementationStatus(StrEnum):
    NOT_IMPLEMENTED = "NOT_IMPLEMENTED"
    EXISTING_DIAGNOSTIC_ONLY = "EXISTING_DIAGNOSTIC_ONLY"
    IMPLEMENTED = "IMPLEMENTED"


class VerificationResult(StrEnum):
    VERIFIED = "VERIFIED"
    REJECTED = "REJECTED"
    INDETERMINATE = "INDETERMINATE"


@dataclass(frozen=True)
class PhaseVerifierContract:
    evidence_kind: EvidenceKind
    verifier_id: str
    contract_version: str
    verification_method: str
    required_scope: tuple[str, ...]
    requires_native_excel: bool
    requires_package_audit: bool = False
    requires_prepared_snapshot_binding: bool = False
    implementation_status: VerifierImplementationStatus = VerifierImplementationStatus.NOT_IMPLEMENTED

    def __post_init__(self) -> None:
        if not isinstance(self.evidence_kind, EvidenceKind):
            raise PhaseVerifierContractError("Contract evidence_kind is invalid")
        for name, value in (("verifier_id", self.verifier_id),
                            ("contract_version", self.contract_version),
                            ("verification_method", self.verification_method)):
            if type(value) is not str or not _ID.fullmatch(value):
                raise PhaseVerifierContractError(f"{name} is invalid")
        if (type(self.required_scope) is not tuple or not self.required_scope
                or any(type(item) is not str or not item.strip() for item in self.required_scope)
                or len(set(self.required_scope)) != len(self.required_scope)):
            raise PhaseVerifierContractError("required_scope must be a unique non-empty tuple")
        if type(self.requires_native_excel) is not bool:
            raise PhaseVerifierContractError("requires_native_excel must be Boolean")
        if type(self.requires_package_audit) is not bool:
            raise PhaseVerifierContractError("requires_package_audit must be Boolean")
        if type(self.requires_prepared_snapshot_binding) is not bool:
            raise PhaseVerifierContractError("requires_prepared_snapshot_binding must be Boolean")
        if not isinstance(self.implementation_status, VerifierImplementationStatus):
            raise PhaseVerifierContractError("implementation_status is invalid")
        if self.requires_package_audit and not self.requires_native_excel:
            raise PhaseVerifierContractError("Package audit contract must also require native Excel reopen")


_CONTRACTS = {
    EvidenceKind.PREPARED_SNAPSHOT: PhaseVerifierContract(
        EvidenceKind.PREPARED_SNAPSHOT, "diomil.prepared-snapshot-native-reopen",
        "contract-v1", "native_excel_readonly_reopen",
        ("saved_workbook_sha256", "daily_input_service_date", "journal_history_prefix",
         "prepared_row_exact_match", "payload_sha256", "journal_sheet_protection_visibility",
         "native_excel_readonly_reopen"), True,
        implementation_status=VerifierImplementationStatus.EXISTING_DIAGNOSTIC_ONLY,
    ),
    EvidenceKind.INPUT_AND_LEDGER_REOPEN: PhaseVerifierContract(
        EvidenceKind.INPUT_AND_LEDGER_REOPEN, "diomil.input-ledger-native-reopen",
        "contract-v1", "native_excel_readonly_reopen",
        ("sheet:DAILY_INPUT", "authoritative_ledger", "operation_identity", "saved_workbook_sha256"), True,
    ),
    EvidenceKind.PROJECTIONS_REOPEN: PhaseVerifierContract(
        EvidenceKind.PROJECTIONS_REOPEN, "diomil.projections-native-reopen",
        "contract-v1", "native_excel_readonly_reopen",
        ("sheet:MASTER_SCHEDULE", "affected_therapist_projections", "source_input_ledger_linkage"), True,
    ),
    EvidenceKind.PENDING_DECISION_INVENTORY: PhaseVerifierContract(
        EvidenceKind.PENDING_DECISION_INVENTORY, "diomil.pending-decision-inventory-reopen",
        "contract-v1", "native_excel_readonly_reopen",
        ("pending_replacement_inventory", "operation_identity", "saved_workbook_sha256"), True,
    ),
    EvidenceKind.REPLACEMENTS_REOPEN: PhaseVerifierContract(
        EvidenceKind.REPLACEMENTS_REOPEN, "diomil.replacements-native-reopen",
        "contract-v1", "native_excel_readonly_reopen",
        ("replacement_facts", "suspensions", "operation_identity", "saved_workbook_sha256"), True,
    ),
    EvidenceKind.OPERATOR_RESOLUTION: PhaseVerifierContract(
        EvidenceKind.OPERATOR_RESOLUTION, "diomil.operator-resolution-record-audit",
        "contract-v1", "independent_operator_decision_record_audit",
        ("operator_decision_record", "operation_id_linkage", "decision_value", "decision_timestamp"), False,
    ),
    EvidenceKind.APPLIED_ARCHIVE_REOPEN: PhaseVerifierContract(
        EvidenceKind.APPLIED_ARCHIVE_REOPEN, "diomil.applied-archive-native-reopen",
        "contract-v1", "native_excel_readonly_reopen",
        ("applied_archive", "operation_identity", "saved_workbook_sha256"), True,
    ),
    EvidenceKind.FINAL_WORKBOOK_REOPEN: PhaseVerifierContract(
        EvidenceKind.FINAL_WORKBOOK_REOPEN, "diomil.final-workbook-native-reopen",
        "contract-v1", "native_excel_readonly_reopen",
        ("full_workbook", "required_invariants", "operation_identity", "saved_workbook_sha256"), True,
    ),
    EvidenceKind.FAILURE_DIAGNOSTIC: PhaseVerifierContract(
        EvidenceKind.FAILURE_DIAGNOSTIC, "diomil.failure-diagnostic-artifact-audit",
        "contract-v1", "operation_bound_failure_artifact_audit",
        ("failure_diagnostic", "operation_identity", "failure_code", "failure_detail", "retained_artifact_sha256"), False,
    ),
    EvidenceKind.FULL_WORKBOOK_ROLLBACK: PhaseVerifierContract(
        EvidenceKind.FULL_WORKBOOK_ROLLBACK, "diomil.full-workbook-rollback-native-reopen",
        "contract-v1", "native_excel_readonly_reopen_plus_package_audit",
        ("restored_workbook", "non_journal_package_comparison", "journal_prefix",
         "prepared_snapshot_identity", "saved_workbook_sha256"), True, True, True,
    ),
}
VERIFIER_CONTRACTS: Mapping[EvidenceKind, PhaseVerifierContract] = MappingProxyType(_CONTRACTS)


def _canonical_json(value: dict) -> str:
    return json.dumps(value, ensure_ascii=True, allow_nan=False, sort_keys=True,
                      separators=(",", ":"))


def _require_text(name: str, value: object, *, identifier: bool = False) -> str:
    if type(value) is not str or not value.strip():
        raise PhaseVerifierContractError(f"{name} must be a non-empty string")
    if identifier and not _ID.fullmatch(value):
        raise PhaseVerifierContractError(f"{name} has invalid identifier syntax")
    return value


def _require_hash(name: str, value: object, *, nullable: bool = False) -> str | None:
    if nullable and value is None:
        return None
    if type(value) is not str or not _HEX.fullmatch(value):
        raise PhaseVerifierContractError(f"{name} must be a lowercase SHA-256 digest")
    return value


def _validate_artifact_locator(value: object) -> str:
    text = _require_text("evidence_artifact", value)
    if "\\" in text or text.startswith("/") or ":" in text:
        raise PhaseVerifierContractError("evidence_artifact must be a relative POSIX locator")
    path = PurePosixPath(text)
    if path.is_absolute() or any(part in {"", ".", ".."} for part in path.parts):
        raise PhaseVerifierContractError("evidence_artifact contains unsafe path segments")
    if path.as_posix() != text:
        raise PhaseVerifierContractError("evidence_artifact is not canonical")
    return text


def _validate_utc(value: object) -> str:
    text = _require_text("verified_at_utc", value)
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as exc:
        raise PhaseVerifierContractError("verified_at_utc must be RFC3339") from exc
    if parsed.tzinfo is None or parsed.utcoffset() != timezone.utc.utcoffset(parsed):
        raise PhaseVerifierContractError("verified_at_utc must be expressed in UTC")
    if not text.endswith("Z") and not text.endswith("+00:00"):
        raise PhaseVerifierContractError("verified_at_utc must use Z or +00:00")
    return text


@dataclass(frozen=True)
class DailyInputPhaseVerifierReport:
    schema: str
    report_id: str
    operation_id: str
    source_workbook_sha256: str
    service_date: str
    journal_revision: int
    journal_phase: str
    evidence_kind: str
    verifier_id: str
    contract_version: str
    verification_method: str
    verification_result: str
    evidence_artifact: str
    evidence_sha256: str
    inspected_scope: tuple[str, ...]
    native_excel_version: str | None
    package_audit_sha256: str | None
    prepared_snapshot_sha256: str | None
    verified_at_utc: str
    report_sha256: str

    def __post_init__(self) -> None:
        if self.schema != PHASE_VERIFIER_REPORT_SCHEMA:
            raise PhaseVerifierContractError("Unsupported phase verifier report schema")
        _require_text("report_id", self.report_id, identifier=True)
        _require_text("operation_id", self.operation_id, identifier=True)
        _require_hash("source_workbook_sha256", self.source_workbook_sha256)
        _require_text("service_date", self.service_date)
        if type(self.journal_revision) is not int or self.journal_revision < 0:
            raise PhaseVerifierContractError("journal_revision must be a non-negative integer")
        _require_text("journal_phase", self.journal_phase)
        try:
            EvidenceKind(self.evidence_kind)
            VerificationResult(self.verification_result)
        except ValueError as exc:
            raise PhaseVerifierContractError("Unknown evidence kind or verification result") from exc
        _require_text("verifier_id", self.verifier_id, identifier=True)
        _require_text("contract_version", self.contract_version, identifier=True)
        _require_text("verification_method", self.verification_method, identifier=True)
        _validate_artifact_locator(self.evidence_artifact)
        _require_hash("evidence_sha256", self.evidence_sha256)
        if (type(self.inspected_scope) is not tuple or not self.inspected_scope
                or any(type(item) is not str or not item.strip() for item in self.inspected_scope)
                or len(set(self.inspected_scope)) != len(self.inspected_scope)):
            raise PhaseVerifierContractError("inspected_scope must be a unique non-empty tuple")
        if self.native_excel_version is not None:
            _require_text("native_excel_version", self.native_excel_version)
        _require_hash("package_audit_sha256", self.package_audit_sha256, nullable=True)
        _require_hash("prepared_snapshot_sha256", self.prepared_snapshot_sha256, nullable=True)
        _validate_utc(self.verified_at_utc)
        _require_hash("report_sha256", self.report_sha256)

    def unsigned_dict(self) -> dict:
        value = asdict(self)
        value.pop("report_sha256")
        value["inspected_scope"] = list(self.inspected_scope)
        return value

    def to_dict(self) -> dict:
        value = asdict(self)
        value["inspected_scope"] = list(self.inspected_scope)
        return value

    def to_json(self) -> str:
        return _canonical_json(self.to_dict())


def create_phase_verifier_report(**fields) -> DailyInputPhaseVerifierReport:
    """Create a canonical report envelope; it does not execute or authenticate a verifier."""
    expected = _REPORT_FIELDS - {"schema", "report_sha256"}
    if set(fields) != expected:
        raise PhaseVerifierContractError("Phase verifier report fields missing or unexpected")
    unsigned = {"schema": PHASE_VERIFIER_REPORT_SCHEMA, **fields}
    digest = sha256(_canonical_json(unsigned).encode("ascii")).hexdigest()
    if type(unsigned.get("inspected_scope")) is list:
        unsigned["inspected_scope"] = tuple(unsigned["inspected_scope"])
    return DailyInputPhaseVerifierReport(**unsigned, report_sha256=digest)


def decode_phase_verifier_report(value: str | bytes | dict) -> DailyInputPhaseVerifierReport:
    """Decode canonical report JSON and verify its self-digest, not its authenticity."""
    try:
        if isinstance(value, bytes):
            value = value.decode("utf-8")
        data = json.loads(value) if isinstance(value, str) else value
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise PhaseVerifierContractError("Phase verifier report JSON is invalid") from exc
    if type(data) is not dict or set(data) != _REPORT_FIELDS:
        raise PhaseVerifierContractError("Phase verifier report fields missing or unexpected")
    normalized = dict(data)
    if type(normalized["inspected_scope"]) is list:
        normalized["inspected_scope"] = tuple(normalized["inspected_scope"])
    report = DailyInputPhaseVerifierReport(**normalized)
    expected = sha256(_canonical_json(report.unsigned_dict()).encode("ascii")).hexdigest()
    if report.report_sha256 != expected:
        raise PhaseVerifierContractError("Phase verifier report SHA-256 mismatch")
    if isinstance(value, str) and _canonical_json(report.to_dict()) != value:
        raise PhaseVerifierContractError("Phase verifier report JSON is not canonical")
    return report


def validate_phase_verifier_report(
    report: DailyInputPhaseVerifierReport,
    operation: DailyInputOperation,
    *,
    previous_operation: DailyInputOperation | None = None,
) -> None:
    """Fail closed unless report matches the exact contract and an implemented verifier.

    Even after implementation status changes, this function only validates report
    structure and binding. Authenticity must come from a separately trusted build
    manifest/attestation mechanism. No report field is accepted as a caller Boolean.
    """
    if type(report) is not DailyInputPhaseVerifierReport or type(operation) is not DailyInputOperation:
        raise PhaseVerifierContractError("Invalid report or operation type")
    expected_digest = sha256(_canonical_json(report.unsigned_dict()).encode("ascii")).hexdigest()
    if report.report_sha256 != expected_digest:
        raise PhaseVerifierContractError("Phase verifier report SHA-256 mismatch")
    try:
        kind = EvidenceKind(report.evidence_kind)
    except ValueError as exc:
        raise PhaseVerifierContractError("Unknown report evidence kind") from exc
    if operation.phase.value == "PREPARED" and operation.revision == 0:
        if previous_operation is not None:
            raise PhaseVerifierContractError("Initial PREPARED report cannot have a previous operation")
        required_evidence = (EvidenceKind.PREPARED_SNAPSHOT,)
    else:
        if previous_operation is None:
            raise PhaseVerifierContractError("Previous operation is required to derive transition evidence")
        try:
            required_evidence = required_evidence_for_transition(previous_operation, operation)
        except EvidencePolicyError as exc:
            raise PhaseVerifierContractError("Operation transition has no valid evidence contract") from exc
    contract = VERIFIER_CONTRACTS.get(kind)
    if contract is None:
        raise PhaseVerifierContractError("No verifier contract exists for evidence kind")
    if kind not in required_evidence:
        raise PhaseVerifierContractError("Report evidence kind is not required for this transition")
    if (report.operation_id, report.source_workbook_sha256, report.service_date,
            report.journal_revision, report.journal_phase) != (
            operation.operation_id, operation.source_sha256, operation.service_date,
            operation.revision, operation.phase.value):
        raise PhaseVerifierContractError("Verifier report does not match operation identity/revision")
    if (report.verifier_id, report.contract_version, report.verification_method) != (
            contract.verifier_id, contract.contract_version, contract.verification_method):
        raise PhaseVerifierContractError("Verifier report does not match pinned contract identity")
    if report.inspected_scope != contract.required_scope:
        raise PhaseVerifierContractError("Verifier report inspected scope differs from contract")
    if contract.requires_native_excel and report.native_excel_version is None:
        raise PhaseVerifierContractError("Native Excel version is required by verifier contract")
    if not contract.requires_native_excel and report.native_excel_version is not None:
        raise PhaseVerifierContractError("Unexpected native Excel metadata for non-Excel verifier")
    if contract.requires_package_audit and report.package_audit_sha256 is None:
        raise PhaseVerifierContractError("Package audit digest is required by verifier contract")
    if not contract.requires_package_audit and report.package_audit_sha256 is not None:
        raise PhaseVerifierContractError("Unexpected package audit digest")
    if contract.requires_prepared_snapshot_binding and report.prepared_snapshot_sha256 is None:
        raise PhaseVerifierContractError("PREPARED snapshot digest is required by verifier contract")
    if not contract.requires_prepared_snapshot_binding and report.prepared_snapshot_sha256 is not None:
        raise PhaseVerifierContractError("Unexpected PREPARED snapshot digest")
    if report.verification_result != VerificationResult.VERIFIED.value:
        raise PhaseVerifierContractError("Only VERIFIED reports can satisfy a positive phase obligation")
    if contract.implementation_status == VerifierImplementationStatus.EXISTING_DIAGNOSTIC_ONLY:
        raise PhaseVerifierNotImplementedError(
            f"Existing diagnostic verifier for {kind.value} is not integrated with the phase report/receipt contract"
        )
    if contract.implementation_status != VerifierImplementationStatus.IMPLEMENTED:
        raise PhaseVerifierNotImplementedError(
            f"No executable independent verifier is registered for {kind.value}"
        )


def registered_verifier_contracts() -> tuple[PhaseVerifierContract, ...]:
    """Return immutable contract metadata; this is not an executable verifier registry."""
    return tuple(VERIFIER_CONTRACTS[kind] for kind in EvidenceKind)
