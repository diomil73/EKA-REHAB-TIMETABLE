"""Read-only proof obligations for DAILY_INPUT journal phase transitions.

This module does NOT verify a workbook, accept caller-supplied proof flags, write
journal rows or enable Apply. It defines the evidence an independent verifier
must establish *before* a durable phase can be appended. No independent phase
verifiers are registered yet: this policy is NOT an authorization gate by itself.
"""
from __future__ import annotations

from enum import StrEnum

from .daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationError,
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)


class PhaseEvidenceContractError(ValueError):
    """Proposed journal transition is not a valid, new operation revision."""


class EvidenceKind(StrEnum):
    PREPARED_SNAPSHOT = "prepared_snapshot_native_readback"
    INPUT_AND_LEDGER_REOPEN = "input_and_ledger_native_reopen"
    PROJECTIONS_REOPEN = "master_and_therapist_projections_native_reopen"
    PENDING_DECISION_INVENTORY = "pending_replacement_inventory_reopen"
    REPLACEMENTS_REOPEN = "replacement_facts_native_reopen"
    OPERATOR_RESOLUTION = "operator_resolution_record_verified"
    APPLIED_ARCHIVE_REOPEN = "applied_archive_native_reopen"
    FINAL_WORKBOOK_REOPEN = "full_workbook_final_native_reopen"
    FAILURE_DIAGNOSTIC = "failure_diagnostic_retained"
    FULL_WORKBOOK_ROLLBACK = "full_workbook_rollback_native_reopen"


_BASE_REQUIREMENTS: dict[Phase, tuple[EvidenceKind, ...]] = {
    Phase.INPUT_AND_LEDGER_COMMITTED: (EvidenceKind.INPUT_AND_LEDGER_REOPEN,),
    Phase.PROJECTIONS_VERIFIED: (EvidenceKind.PROJECTIONS_REOPEN,),
    Phase.PENDING_OPERATOR_DECISION: (EvidenceKind.PENDING_DECISION_INVENTORY,),
    Phase.REPLACEMENTS_RESOLVED: (EvidenceKind.REPLACEMENTS_REOPEN,),
    Phase.APPLIED_ARCHIVE_VERIFIED: (EvidenceKind.APPLIED_ARCHIVE_REOPEN,),
    Phase.COMPLETE: (EvidenceKind.FINAL_WORKBOOK_REOPEN,),
    Phase.FAILED: (EvidenceKind.FAILURE_DIAGNOSTIC,),
    Phase.ROLLED_BACK: (EvidenceKind.FULL_WORKBOOK_ROLLBACK,),
}


def required_evidence_for_transition(
    previous: DailyInputOperation,
    candidate: DailyInputOperation,
) -> tuple[EvidenceKind, ...]:
    """Check legal next-state identity and return independent proof obligations.

    The return value is a *requirements list*, NEVER a verified receipt or
    permission to append. Each proof kind still needs a separate trusted native
    workbook/readback verifier and source-bound handoff. No caller Boolean or
    untrusted metadata can satisfy these obligations.
    """
    if type(previous) is not DailyInputOperation or type(candidate) is not DailyInputOperation:
        raise PhaseEvidenceContractError("Invalid transition operation type")
    if candidate.phase in {Phase.PREPARED, previous.phase}:
        raise PhaseEvidenceContractError("A new non-PREPARED journal revision is required")
    if (
        candidate.operation_id != previous.operation_id
        or candidate.source_sha256 != previous.source_sha256
        or candidate.service_date != previous.service_date
        or candidate.revision != previous.revision + 1
    ):
        raise PhaseEvidenceContractError("Operation identity or revision changed")

    kwargs: dict[str, object] = {}
    if candidate.phase == Phase.FAILED:
        kwargs["failure_code"] = candidate.failure_code
        kwargs["failure_detail"] = candidate.failure_detail
    elif candidate.phase == Phase.ROLLED_BACK:
        kwargs["rollback_verified"] = candidate.rollback_verified
    elif previous.phase == Phase.PENDING_OPERATOR_DECISION and candidate.phase == Phase.REPLACEMENTS_RESOLVED:
        kwargs["operator_resolution"] = candidate.operator_resolution

    try:
        expected = advance_daily_input_operation(previous, candidate.phase, **kwargs)
    except DailyInputOperationError as exc:
        raise PhaseEvidenceContractError("Illegal operation phase transition") from exc
    if expected != candidate:
        raise PhaseEvidenceContractError("Transition metadata does not match legal next state")

    required = _BASE_REQUIREMENTS.get(candidate.phase)
    if not required:
        raise PhaseEvidenceContractError("No proof contract for target phase")
    if (
        previous.phase == Phase.PENDING_OPERATOR_DECISION
        and candidate.phase == Phase.REPLACEMENTS_RESOLVED
    ):
        return (EvidenceKind.OPERATOR_RESOLUTION, *required)
    return required
