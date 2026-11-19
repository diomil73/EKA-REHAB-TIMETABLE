"""Explicit lifecycle contract for one DAILY_INPUT Apply operation.

This module is deliberately pure: it defines legal state transitions and replay
semantics. It does not persist a journal or claim that Excel side effects are
atomic. The native VBA orchestrator must persist each returned state and retain
its rollback evidence until COMPLETE or verified ROLLED_BACK.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from enum import StrEnum


class DailyInputOperationError(ValueError):
    """Raised when an operation attempts an illegal lifecycle transition."""


class DailyInputOperationPhase(StrEnum):
    PREPARED = "PREPARED"
    INPUT_AND_LEDGER_COMMITTED = "INPUT_AND_LEDGER_COMMITTED"
    PROJECTIONS_VERIFIED = "PROJECTIONS_VERIFIED"
    PENDING_OPERATOR_DECISION = "PENDING_OPERATOR_DECISION"
    REPLACEMENTS_RESOLVED = "REPLACEMENTS_RESOLVED"
    APPLIED_ARCHIVE_VERIFIED = "APPLIED_ARCHIVE_VERIFIED"
    COMPLETE = "COMPLETE"
    FAILED = "FAILED"
    ROLLED_BACK = "ROLLED_BACK"


_ALLOWED: dict[DailyInputOperationPhase, frozenset[DailyInputOperationPhase]] = {
    DailyInputOperationPhase.PREPARED: frozenset({
        DailyInputOperationPhase.INPUT_AND_LEDGER_COMMITTED,
        DailyInputOperationPhase.FAILED,
    }),
    DailyInputOperationPhase.INPUT_AND_LEDGER_COMMITTED: frozenset({
        DailyInputOperationPhase.PROJECTIONS_VERIFIED,
        DailyInputOperationPhase.FAILED,
    }),
    DailyInputOperationPhase.PROJECTIONS_VERIFIED: frozenset({
        DailyInputOperationPhase.PENDING_OPERATOR_DECISION,
        DailyInputOperationPhase.REPLACEMENTS_RESOLVED,
        DailyInputOperationPhase.FAILED,
    }),
    DailyInputOperationPhase.PENDING_OPERATOR_DECISION: frozenset({
        DailyInputOperationPhase.REPLACEMENTS_RESOLVED,
        DailyInputOperationPhase.FAILED,
    }),
    DailyInputOperationPhase.REPLACEMENTS_RESOLVED: frozenset({
        DailyInputOperationPhase.APPLIED_ARCHIVE_VERIFIED,
        DailyInputOperationPhase.FAILED,
    }),
    DailyInputOperationPhase.APPLIED_ARCHIVE_VERIFIED: frozenset({
        DailyInputOperationPhase.COMPLETE,
        DailyInputOperationPhase.FAILED,
    }),
    DailyInputOperationPhase.FAILED: frozenset({
        DailyInputOperationPhase.ROLLED_BACK,
    }),
    DailyInputOperationPhase.COMPLETE: frozenset(),
    DailyInputOperationPhase.ROLLED_BACK: frozenset(),
}


@dataclass(frozen=True)
class DailyInputOperation:
    operation_id: str
    source_sha256: str
    service_date: str
    phase: DailyInputOperationPhase = DailyInputOperationPhase.PREPARED
    revision: int = 0
    failure_code: str = ""
    failure_detail: str = ""
    operator_resolution: str = ""

    def __post_init__(self) -> None:
        if not self.operation_id.strip():
            raise DailyInputOperationError("operation_id must be non-empty")
        if len(self.source_sha256) != 64 or any(
            char not in "0123456789abcdef" for char in self.source_sha256.lower()
        ):
            raise DailyInputOperationError("source_sha256 must be a 64-character hex digest")
        if not self.service_date.strip():
            raise DailyInputOperationError("service_date must be non-empty")
        if self.revision < 0:
            raise DailyInputOperationError("revision cannot be negative")
        if self.phase == DailyInputOperationPhase.FAILED and not self.failure_code.strip():
            raise DailyInputOperationError("FAILED requires a failure_code")
        if self.phase == DailyInputOperationPhase.ROLLED_BACK and not self.failure_code.strip():
            raise DailyInputOperationError("ROLLED_BACK requires retained failure evidence")
        if self.phase == DailyInputOperationPhase.PENDING_OPERATOR_DECISION and self.operator_resolution:
            raise DailyInputOperationError("pending decision cannot already have a resolution")
        if self.operator_resolution and self.phase not in {
            DailyInputOperationPhase.REPLACEMENTS_RESOLVED,
            DailyInputOperationPhase.APPLIED_ARCHIVE_VERIFIED,
            DailyInputOperationPhase.COMPLETE,
        }:
            raise DailyInputOperationError("operator resolution is recorded before replacement resolution")

    @property
    def is_terminal(self) -> bool:
        return self.phase in {
            DailyInputOperationPhase.COMPLETE,
            DailyInputOperationPhase.ROLLED_BACK,
        }

    @property
    def may_report_success(self) -> bool:
        return self.phase == DailyInputOperationPhase.COMPLETE


def advance_daily_input_operation(
    operation: DailyInputOperation,
    target: DailyInputOperationPhase,
    *,
    operator_resolution: str = "",
    failure_code: str = "",
    failure_detail: str = "",
) -> DailyInputOperation:
    """Return the next state; exact retries are idempotent, conflicting retries fail.

    Allowed operator resolutions are 'automatic_accepted' and 'manual_completed'.
    A declined automatic proposal is not a terminal outcome: it remains pending
    until the manual decision is completed or the full operation is rolled back.
    """
    if target == operation.phase:
        if (
            operator_resolution == operation.operator_resolution
            and failure_code == operation.failure_code
            and failure_detail == operation.failure_detail
        ):
            return operation
        raise DailyInputOperationError("Conflicting replay for the current operation phase")
    if operation.is_terminal:
        raise DailyInputOperationError("Terminal operation cannot transition")
    if target not in _ALLOWED[operation.phase]:
        raise DailyInputOperationError(
            f"Illegal DAILY_INPUT transition: {operation.phase.value} -> {target.value}"
        )

    changes: dict[str, object] = {"phase": target, "revision": operation.revision + 1}
    if target == DailyInputOperationPhase.PENDING_OPERATOR_DECISION:
        if operator_resolution:
            raise DailyInputOperationError("A pending decision cannot include a resolution")
    elif target == DailyInputOperationPhase.REPLACEMENTS_RESOLVED:
        if operation.phase == DailyInputOperationPhase.PENDING_OPERATOR_DECISION:
            if operator_resolution not in {"automatic_accepted", "manual_completed"}:
                raise DailyInputOperationError(
                    "Operator decision must be automatic_accepted or manual_completed"
                )
            changes["operator_resolution"] = operator_resolution
        elif operator_resolution:
            raise DailyInputOperationError("Unexpected operator resolution without pending review")
    elif target == DailyInputOperationPhase.FAILED:
        if not failure_code.strip():
            raise DailyInputOperationError("FAILED transition requires failure_code")
        changes["failure_code"] = failure_code.strip()
        changes["failure_detail"] = failure_detail.strip()
    elif target == DailyInputOperationPhase.ROLLED_BACK:
        if operation.phase != DailyInputOperationPhase.FAILED:
            raise DailyInputOperationError("Rollback is only valid after FAILED")
    elif operator_resolution or failure_code or failure_detail:
        raise DailyInputOperationError("Unexpected transition metadata")

    return replace(operation, **changes)
