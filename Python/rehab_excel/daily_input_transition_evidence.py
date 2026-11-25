"""Structural completeness check for one DAILY_INPUT phase transition.

This module checks that every lifecycle-derived evidence obligation has exactly
one matching report and receipt, that all records bind to the same transition,
and that receipt artifacts form one exact byte-verified predecessor chain.

It does NOT authenticate verifier origin, validate a trusted build attestation,
append a journal row, authorize Apply/Restore, or replace native Excel verifiers.
Current verifier registry entries intentionally remain non-authorizing until
their independent implementations and trusted attestation are integrated.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .daily_input_operation import DailyInputOperation
from .daily_input_phase_evidence import (
    EvidenceKind,
    PhaseEvidenceContractError,
    required_evidence_for_transition,
)
from .daily_input_phase_verifier_contract import (
    DailyInputPhaseVerifierReport,
    validate_phase_verifier_report,
)
from .daily_input_transaction_receipt import (
    DailyInputTransactionReceipt,
    validate_receipt_binding,
    verify_receipt_artifact_bytes,
)


class TransitionEvidenceBundleError(ValueError):
    """The evidence bundle is incomplete, duplicated, inconsistent or unbound."""


@dataclass(frozen=True)
class StructurallyValidatedTransitionEvidenceBundle:
    """Immutable structural result; deliberately not an authorization token."""

    operation_id: str
    journal_revision: int
    journal_phase: str
    required_evidence: tuple[EvidenceKind, ...]
    reports: tuple[DailyInputPhaseVerifierReport, ...]
    receipts: tuple[DailyInputTransactionReceipt, ...]


def _index_exactly_once(items, *, item_type, label: str) -> dict[EvidenceKind, object]:
    if type(items) is not tuple:
        raise TransitionEvidenceBundleError(f"{label} must be supplied as an immutable tuple")
    indexed: dict[EvidenceKind, object] = {}
    for item in items:
        if type(item) is not item_type:
            raise TransitionEvidenceBundleError(f"Invalid {label} item type")
        try:
            kind = EvidenceKind(item.evidence_kind)
        except ValueError as exc:
            raise TransitionEvidenceBundleError(f"Unknown {label} evidence kind") from exc
        if kind in indexed:
            raise TransitionEvidenceBundleError(f"Duplicate {label} for evidence kind {kind.value}")
        indexed[kind] = item
    return indexed


def validate_transition_evidence_bundle(
    previous_operation: DailyInputOperation,
    operation: DailyInputOperation,
    *,
    reports: tuple[DailyInputPhaseVerifierReport, ...],
    receipts: tuple[DailyInputTransactionReceipt, ...],
    expected_previous_receipt_sha256: str,
    verifier_registry: dict[str, str],
    evidence_root: str | Path,
) -> StructurallyValidatedTransitionEvidenceBundle:
    """Require the exact complete evidence set and a byte-verified receipt chain.

    Required evidence is always derived from the previous/candidate lifecycle
    states. Callers cannot provide or narrow the obligation list. A successful
    return means structural consistency only: verifier authenticity and phase
    authorization remain separate, unavailable controls.
    """
    if type(previous_operation) is not DailyInputOperation or type(operation) is not DailyInputOperation:
        raise TransitionEvidenceBundleError("Invalid previous or candidate operation")
    try:
        required = required_evidence_for_transition(previous_operation, operation)
    except PhaseEvidenceContractError as exc:
        raise TransitionEvidenceBundleError("Invalid lifecycle transition") from exc

    report_by_kind = _index_exactly_once(
        reports, item_type=DailyInputPhaseVerifierReport, label="phase verifier report"
    )
    receipt_by_kind = _index_exactly_once(
        receipts, item_type=DailyInputTransactionReceipt, label="transaction receipt"
    )
    expected = set(required)
    if set(report_by_kind) != expected:
        missing = expected - set(report_by_kind)
        extra = set(report_by_kind) - expected
        raise TransitionEvidenceBundleError(
            f"Report evidence set is incomplete or excessive; missing={sorted(k.value for k in missing)}; "
            f"extra={sorted(k.value for k in extra)}"
        )
    if set(receipt_by_kind) != expected:
        missing = expected - set(receipt_by_kind)
        extra = set(receipt_by_kind) - expected
        raise TransitionEvidenceBundleError(
            f"Receipt evidence set is incomplete or excessive; missing={sorted(k.value for k in missing)}; "
            f"extra={sorted(k.value for k in extra)}"
        )

    ordered_reports = tuple(report_by_kind[kind] for kind in required)
    ordered_receipts = tuple(receipt_by_kind[kind] for kind in required)
    predecessor = expected_previous_receipt_sha256

    for report, receipt in zip(ordered_reports, ordered_receipts, strict=True):
        try:
            validate_phase_verifier_report(
                report, operation, previous_operation=previous_operation
            )
            if (
                report.evidence_kind != receipt.evidence_kind
                or report.evidence_artifact != receipt.evidence_artifact
                or report.evidence_sha256 != receipt.evidence_sha256
                or report.verifier_id != receipt.verifier_id
                or report.contract_version != receipt.verifier_version
                or report.verification_method != receipt.verification_method
                or report.verification_result != receipt.verification_result
            ):
                raise TransitionEvidenceBundleError(
                    "Report and receipt do not describe the same evidence artifact/verifier"
                )
            validate_receipt_binding(
                receipt, operation,
                previous_operation=previous_operation,
                expected_previous_receipt_sha256=predecessor,
                verifier_registry=verifier_registry,
            )
            verify_receipt_artifact_bytes(receipt, evidence_root)
        except TransitionEvidenceBundleError:
            raise
        except (ValueError, OSError) as exc:
            raise TransitionEvidenceBundleError(
                f"Evidence validation failed for {report.evidence_kind}"
            ) from exc
        predecessor = receipt.receipt_sha256

    return StructurallyValidatedTransitionEvidenceBundle(
        operation_id=operation.operation_id,
        journal_revision=operation.revision,
        journal_phase=operation.phase.value,
        required_evidence=required,
        reports=ordered_reports,
        receipts=ordered_receipts,
    )
