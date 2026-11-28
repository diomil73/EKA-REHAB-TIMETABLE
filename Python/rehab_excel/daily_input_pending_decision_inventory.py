"""Canonical, read-only snapshot of unresolved therapist-replacement targets.

This is an inventory builder, not a replacement chooser or native Excel verifier.
It derives pending targets only from the already validated projection preflight
state model, binds them to the operation and exact preflight workbook digest,
and can publish an exclusive local evidence artifact. It does not include
provider options, make decisions, write a workbook, create a receipt, or
authorize a lifecycle transition.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date, time
from hashlib import sha256
import json
import os
from pathlib import Path
import re

from rehab_core.daily_state import DailySessionStatus, DailySessionState
from rehab_core.models import Session

from .daily_input_operation import DailyInputOperation, DailyInputOperationPhase
from .daily_input_phase_evidence import (
    EvidenceKind,
    PhaseEvidenceContractError,
    required_evidence_for_transition,
)
from .daily_input_projection_preflight import DailyInputProjectionPreflight


SCHEMA = "diomil-pending-decision-inventory-v1"
_HEX = re.compile(r"^[0-9a-f]{64}$")


class PendingDecisionInventoryError(ValueError):
    """Pending replacement inventory is incomplete, inconsistent or unsafe."""


@dataclass(frozen=True)
class PendingReplacementTarget:
    session_id: str
    patient_id: str
    original_therapist_id: str
    original_time: str
    treatment: str | None
    robotic: bool


@dataclass(frozen=True)
class PendingDecisionInventory:
    schema: str
    operation_id: str
    operation_source_sha256: str
    service_date: str
    evidence_workbook_sha256: str
    total_session_count: int
    active_count: int
    patient_absent_count: int
    therapist_absent_count: int
    replaced_count: int
    cancelled_count: int
    targets: tuple[PendingReplacementTarget, ...]
    inventory_sha256: str

    def unsigned_dict(self) -> dict:
        value = asdict(self)
        value.pop("inventory_sha256")
        value["targets"] = [asdict(item) for item in self.targets]
        return value

    def to_dict(self) -> dict:
        value = self.unsigned_dict()
        value["inventory_sha256"] = self.inventory_sha256
        return value

    def to_json(self) -> str:
        return json.dumps(
            self.to_dict(), ensure_ascii=True, allow_nan=False,
            sort_keys=True, separators=(",", ":"),
        )


@dataclass(frozen=True)
class PendingDecisionInventoryArtifact:
    artifact_locator: str
    artifact_sha256: str
    byte_count: int


def _canonical_json(value: dict) -> str:
    return json.dumps(
        value, ensure_ascii=True, allow_nan=False,
        sort_keys=True, separators=(",", ":"),
    )


def build_pending_decision_inventory(
    previous_operation: DailyInputOperation,
    operation: DailyInputOperation,
    preflight: DailyInputProjectionPreflight,
) -> PendingDecisionInventory:
    """Derive every unresolved target from the exact dated preflight states.

    A PENDING_OPERATOR_DECISION transition is legal only after
    PROJECTIONS_VERIFIED. Its inventory must contain every and only session
    whose independently built DailySessionState still needs replacement.
    Patient-absent, cancelled, already replaced and active sessions are counted
    separately and are never included as pending targets.
    """
    if (type(previous_operation) is not DailyInputOperation
            or type(operation) is not DailyInputOperation
            or type(preflight) is not DailyInputProjectionPreflight):
        raise PendingDecisionInventoryError("Exact operation and projection preflight types required")
    try:
        obligations = required_evidence_for_transition(previous_operation, operation)
    except PhaseEvidenceContractError as exc:
        raise PendingDecisionInventoryError("Invalid operation phase transition") from exc
    if (operation.phase != DailyInputOperationPhase.PENDING_OPERATOR_DECISION
            or obligations != (EvidenceKind.PENDING_DECISION_INVENTORY,)):
        raise PendingDecisionInventoryError("Expected PENDING_OPERATOR_DECISION transition")
    if (type(preflight.service_date) is not date
            or preflight.service_date.isoformat() != operation.service_date):
        raise PendingDecisionInventoryError("Preflight service date differs from operation")
    if (type(operation.source_sha256) is not str
            or not _HEX.fullmatch(operation.source_sha256)
            or type(preflight.source_workbook_sha256) is not str
            or not _HEX.fullmatch(preflight.source_workbook_sha256)):
        raise PendingDecisionInventoryError("Operation or evidence workbook digest is invalid")
    if type(preflight.sessions) is not tuple or type(preflight.states) is not tuple:
        raise PendingDecisionInventoryError("Immutable dated sessions and states are required")

    sessions_by_id: dict[str, Session] = {}
    for session in preflight.sessions:
        if type(session) is not Session:
            raise PendingDecisionInventoryError("Invalid preflight session type")
        if (session.session_date != preflight.service_date
                or type(session.session_id) is not str or not session.session_id.strip()
                or type(session.patient_id) is not str or not session.patient_id.strip()
                or type(session.therapist_id) is not str or not session.therapist_id.strip()
                or type(session.start_time) is not time
                or session.start_time.second or session.start_time.microsecond
                or (session.treatment is not None and type(session.treatment) is not str)
                or type(session.robotic) is not bool):
            raise PendingDecisionInventoryError("Malformed or wrong-date preflight session")
        if session.session_id in sessions_by_id:
            raise PendingDecisionInventoryError("Duplicate preflight session identity")
        sessions_by_id[session.session_id] = session
    if not sessions_by_id:
        raise PendingDecisionInventoryError("Pending decision inventory has no dated sessions")

    states_by_id: dict[str, DailySessionState] = {}
    for state in preflight.states:
        if type(state) is not DailySessionState:
            raise PendingDecisionInventoryError("Invalid preflight daily-state type")
        if type(state.session_id) is not str or not state.session_id.strip():
            raise PendingDecisionInventoryError("Invalid preflight daily-state identity")
        if type(state.status) is not DailySessionStatus:
            raise PendingDecisionInventoryError("Invalid preflight daily-state status")
        if state.session_id in states_by_id:
            raise PendingDecisionInventoryError("Duplicate preflight daily-state identity")
        states_by_id[state.session_id] = state
    if set(states_by_id) != set(sessions_by_id):
        raise PendingDecisionInventoryError("Daily states do not cover the exact dated session set")

    counts = {status: 0 for status in DailySessionStatus}
    targets: list[PendingReplacementTarget] = []
    for session_id, session in sessions_by_id.items():
        state = states_by_id[session_id]
        if (state.patient_id != session.patient_id
                or state.session_date != session.session_date
                or state.original_therapist_id != session.therapist_id
                or state.original_time != session.start_time):
            raise PendingDecisionInventoryError("Daily state identity differs from its base session")
        counts[state.status] += 1
        if state.needs_replacement:
            if (state.effective_therapist_id is not None
                    or state.effective_time is not None
                    or state.replacement_id is not None):
                raise PendingDecisionInventoryError("Pending target already contains replacement outcome")
            targets.append(PendingReplacementTarget(
                session_id=session.session_id,
                patient_id=session.patient_id,
                original_therapist_id=session.therapist_id,
                original_time=session.start_time.strftime("%H:%M"),
                treatment=session.treatment,
                robotic=session.robotic,
            ))

    if not targets:
        raise PendingDecisionInventoryError("PENDING_OPERATOR_DECISION requires at least one unresolved target")
    targets.sort(key=lambda item: (
        item.original_time, item.original_therapist_id.casefold(), item.session_id,
    ))
    unsigned = {
        "schema": SCHEMA,
        "operation_id": operation.operation_id,
        "operation_source_sha256": operation.source_sha256,
        "service_date": operation.service_date,
        "evidence_workbook_sha256": preflight.source_workbook_sha256,
        "total_session_count": len(sessions_by_id),
        "active_count": counts[DailySessionStatus.ACTIVE],
        "patient_absent_count": counts[DailySessionStatus.PATIENT_ABSENT],
        "therapist_absent_count": counts[DailySessionStatus.THERAPIST_ABSENT],
        "replaced_count": counts[DailySessionStatus.REPLACED],
        "cancelled_count": counts[DailySessionStatus.CANCELLED],
        "targets": [asdict(item) for item in targets],
    }
    digest = sha256(_canonical_json(unsigned).encode("ascii")).hexdigest()
    return PendingDecisionInventory(
        schema=SCHEMA,
        operation_id=operation.operation_id,
        operation_source_sha256=operation.source_sha256,
        service_date=operation.service_date,
        evidence_workbook_sha256=preflight.source_workbook_sha256,
        total_session_count=len(sessions_by_id),
        active_count=counts[DailySessionStatus.ACTIVE],
        patient_absent_count=counts[DailySessionStatus.PATIENT_ABSENT],
        therapist_absent_count=counts[DailySessionStatus.THERAPIST_ABSENT],
        replaced_count=counts[DailySessionStatus.REPLACED],
        cancelled_count=counts[DailySessionStatus.CANCELLED],
        targets=tuple(targets),
        inventory_sha256=digest,
    )


def validate_pending_decision_inventory(inventory: PendingDecisionInventory) -> None:
    """Revalidate a possibly caller-constructed inventory before persistence."""
    if type(inventory) is not PendingDecisionInventory or inventory.schema != SCHEMA:
        raise PendingDecisionInventoryError("Exact supported pending inventory required")
    for value in (inventory.operation_source_sha256, inventory.evidence_workbook_sha256,
                  inventory.inventory_sha256):
        if type(value) is not str or not _HEX.fullmatch(value):
            raise PendingDecisionInventoryError("Pending inventory digest is invalid")
    if (type(inventory.operation_id) is not str or not inventory.operation_id.strip()
            or type(inventory.service_date) is not str or not inventory.service_date.strip()):
        raise PendingDecisionInventoryError("Pending inventory identity is invalid")
    try:
        from datetime import date
        if date.fromisoformat(inventory.service_date).isoformat() != inventory.service_date:
            raise ValueError("non-canonical date")
    except ValueError as exc:
        raise PendingDecisionInventoryError("Pending inventory service date is invalid") from exc
    counts = (
        inventory.active_count, inventory.patient_absent_count,
        inventory.therapist_absent_count, inventory.replaced_count,
        inventory.cancelled_count,
    )
    if (any(type(value) is not int or value < 0 for value in counts)
            or type(inventory.total_session_count) is not int
            or inventory.total_session_count != sum(counts)
            or type(inventory.targets) is not tuple
            or len(inventory.targets) != inventory.therapist_absent_count
            or not inventory.targets):
        raise PendingDecisionInventoryError("Pending inventory counts or target set are inconsistent")
    seen: set[str] = set()
    for target in inventory.targets:
        if (type(target) is not PendingReplacementTarget
                or type(target.session_id) is not str or not target.session_id.strip()
                or target.session_id in seen
                or type(target.patient_id) is not str or not target.patient_id.strip()
                or type(target.original_therapist_id) is not str or not target.original_therapist_id.strip()
                or type(target.original_time) is not str
                or not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", target.original_time)
                or (target.treatment is not None and type(target.treatment) is not str)
                or type(target.robotic) is not bool):
            raise PendingDecisionInventoryError("Pending inventory target is malformed or duplicated")
        seen.add(target.session_id)
    expected = sha256(_canonical_json(inventory.unsigned_dict()).encode("ascii")).hexdigest()
    if inventory.inventory_sha256 != expected:
        raise PendingDecisionInventoryError("Pending inventory digest mismatch")


def write_pending_decision_inventory(
    inventory: PendingDecisionInventory,
    *,
    evidence_root: str | Path,
    artifact_path: str | Path,
) -> PendingDecisionInventoryArtifact:
    """Write canonical inventory bytes once, exclusively, under evidence_root."""
    validate_pending_decision_inventory(inventory)
    try:
        root = Path(evidence_root).resolve(strict=True)
        parent = Path(artifact_path).parent.resolve(strict=True)
        name = Path(artifact_path).name
        if not root.is_dir() or not name or name in {".", ".."}:
            raise PendingDecisionInventoryError("Invalid evidence root or artifact filename")
        locator_parent = parent.relative_to(root)
        locator = (locator_parent / name).as_posix()
        if locator in {"", "."} or any(part in {"", ".", ".."} for part in Path(locator).parts):
            raise PendingDecisionInventoryError("Unsafe evidence artifact locator")
        output = parent / name
    except PendingDecisionInventoryError:
        raise
    except (OSError, ValueError) as exc:
        raise PendingDecisionInventoryError("Evidence artifact path escapes evidence root") from exc

    payload = inventory.to_json().encode("utf-8")
    digest = sha256(payload).hexdigest()
    created = False
    try:
        with output.open("xb") as handle:
            created = True
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
    except FileExistsError as exc:
        raise PendingDecisionInventoryError("Evidence artifact already exists") from exc
    except OSError as exc:
        if created:
            output.unlink(missing_ok=True)
        raise PendingDecisionInventoryError("Could not persist pending inventory artifact") from exc
    return PendingDecisionInventoryArtifact(locator, digest, len(payload))
