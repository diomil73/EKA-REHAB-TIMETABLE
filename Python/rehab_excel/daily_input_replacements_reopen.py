"""Canonical evidence for resolved daily replacement and treatment-suspension facts.

This module compares the final dated preflight against the exact pending-target
inventory from the preceding phase. Every pending target must resolve to one
and only one active replacement fact or treatment-suspension cancellation.
It creates evidence only; it does not choose providers, write workbooks, append
journal rows, or authorize a lifecycle transition.
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
from rehab_core.models import (
    DailySessionCancellation, ReplacementAssignment, Session,
    SessionCancellationKind,
)

from .daily_input_operation import DailyInputOperation, DailyInputOperationPhase
from .daily_input_pending_decision_inventory import (
    PendingDecisionInventory, PendingDecisionInventoryError,
    read_pending_decision_inventory_artifact, validate_pending_decision_inventory,
)
from .daily_input_phase_evidence import (
    EvidenceKind, PhaseEvidenceContractError, required_evidence_for_transition,
)
from .daily_input_projection_preflight import DailyInputProjectionPreflight
from .daily_replacement_facts import (
    DailyReplacementFactError, DailyReplacementFactRow,
    build_daily_replacement_fact_rows,
)

SCHEMA = "diomil-replacements-reopen-v1"
_HEX = re.compile(r"^[0-9a-f]{64}$")


class ReplacementsReopenError(ValueError):
    """Replacement/suspension evidence is incomplete, inconsistent or unsafe."""


@dataclass(frozen=True)
class ReopenedReplacement:
    replacement_id: str
    target_session_id: str
    patient_id: str
    original_therapist_id: str
    replacement_therapist_id: str
    replacement_date: str
    replacement_time: str
    provider_kind: str


@dataclass(frozen=True)
class ReopenedSuspension:
    cancellation_id: str
    target_session_id: str
    patient_id: str
    original_therapist_id: str
    cancellation_date: str
    cancellation_kind: str


@dataclass(frozen=True)
class ReplacementsReopenEvidence:
    schema: str
    operation_id: str
    operation_source_sha256: str
    service_date: str
    prior_inventory_sha256: str
    prior_inventory_workbook_sha256: str
    evidence_workbook_sha256: str
    pending_target_count: int
    pending_replaced_count: int
    pending_suspended_count: int
    replacement_count: int
    suspension_count: int
    replacements: tuple[ReopenedReplacement, ...]
    suspensions: tuple[ReopenedSuspension, ...]
    evidence_sha256: str

    def unsigned_dict(self) -> dict:
        value = asdict(self)
        value.pop("evidence_sha256")
        value["replacements"] = [asdict(item) for item in self.replacements]
        value["suspensions"] = [asdict(item) for item in self.suspensions]
        return value

    def to_dict(self) -> dict:
        value = self.unsigned_dict()
        value["evidence_sha256"] = self.evidence_sha256
        return value

    def to_json(self) -> str:
        return _canonical_json(self.to_dict())


@dataclass(frozen=True)
class ReplacementsReopenArtifact:
    artifact_locator: str
    artifact_sha256: str
    byte_count: int


def _canonical_json(value: dict) -> str:
    return json.dumps(value, ensure_ascii=True, allow_nan=False,
                      sort_keys=True, separators=(",", ":"))


def _replacement_payload(row: DailyReplacementFactRow) -> ReopenedReplacement:
    return ReopenedReplacement(
        replacement_id=row.replacement_id,
        target_session_id=row.target_session_id,
        patient_id=row.patient_id,
        original_therapist_id=row.original_therapist_id,
        replacement_therapist_id=row.replacement_therapist_id,
        replacement_date=row.replacement_date.isoformat(),
        replacement_time=row.replacement_time.strftime("%H:%M"),
        provider_kind=row.provider_kind.value,
    )


def build_replacements_reopen_evidence(
    previous_operation: DailyInputOperation,
    operation: DailyInputOperation,
    preflight: DailyInputProjectionPreflight,
    pending_inventory: PendingDecisionInventory,
) -> ReplacementsReopenEvidence:
    """Bind the final replacement/suspension partition to prior pending targets."""
    if (type(previous_operation) is not DailyInputOperation
            or type(operation) is not DailyInputOperation
            or type(preflight) is not DailyInputProjectionPreflight
            or type(pending_inventory) is not PendingDecisionInventory):
        raise ReplacementsReopenError("Exact operation, preflight and pending inventory types required")
    try:
        obligations = required_evidence_for_transition(previous_operation, operation)
        validate_pending_decision_inventory(pending_inventory)
    except (PhaseEvidenceContractError, PendingDecisionInventoryError) as exc:
        raise ReplacementsReopenError("Operation transition or pending inventory is invalid") from exc
    if (previous_operation.phase != DailyInputOperationPhase.PENDING_OPERATOR_DECISION
            or operation.phase != DailyInputOperationPhase.REPLACEMENTS_RESOLVED
            or obligations != (EvidenceKind.OPERATOR_RESOLUTION, EvidenceKind.REPLACEMENTS_REOPEN)
            or operation.operator_resolution not in {"automatic_accepted", "manual_completed"}):
        raise ReplacementsReopenError("Expected pending-decision to replacements-resolved transition")
    if (pending_inventory.operation_id != previous_operation.operation_id
            or pending_inventory.operation_source_sha256 != previous_operation.source_sha256
            or pending_inventory.service_date != previous_operation.service_date
            or pending_inventory.evidence_workbook_sha256 == ""):
        raise ReplacementsReopenError("Prior inventory does not bind to the preceding operation")
    if (operation.operation_id != pending_inventory.operation_id
            or operation.source_sha256 != pending_inventory.operation_source_sha256
            or operation.service_date != pending_inventory.service_date
            or type(preflight.service_date) is not date
            or preflight.service_date.isoformat() != operation.service_date
            or type(preflight.source_workbook_sha256) is not str
            or not _HEX.fullmatch(preflight.source_workbook_sha256)):
        raise ReplacementsReopenError("Final preflight does not bind to operation identity/date")

    sessions: dict[str, Session] = {}
    for session in preflight.sessions:
        if (type(session) is not Session or session.session_date != preflight.service_date
                or type(session.session_id) is not str or not session.session_id.strip()
                or type(session.patient_id) is not str or not session.patient_id.strip()
                or type(session.therapist_id) is not str or not session.therapist_id.strip()
                or type(session.start_time) is not time
                or session.start_time.second or session.start_time.microsecond):
            raise ReplacementsReopenError("Malformed or wrong-date final session")
        if session.session_id in sessions:
            raise ReplacementsReopenError("Duplicate final session identity")
        sessions[session.session_id] = session
    if not sessions:
        raise ReplacementsReopenError("Final preflight has no dated sessions")

    states: dict[str, DailySessionState] = {}
    for state in preflight.states:
        if (type(state) is not DailySessionState
                or type(state.status) is not DailySessionStatus
                or state.session_id in states):
            raise ReplacementsReopenError("Malformed or duplicate final daily state")
        states[state.session_id] = state
    if set(states) != set(sessions):
        raise ReplacementsReopenError("Final daily states do not cover exact dated sessions")
    for session_id, session in sessions.items():
        state = states[session_id]
        if (state.patient_id != session.patient_id
                or state.session_date != session.session_date
                or state.original_therapist_id != session.therapist_id
                or state.original_time != session.start_time):
            raise ReplacementsReopenError("Final daily state differs from immutable base session")

    pending = {item.session_id: item for item in pending_inventory.targets}
    if len(pending) != pending_inventory.therapist_absent_count:
        raise ReplacementsReopenError("Pending inventory target count is inconsistent")
    for session_id, target in pending.items():
        session = sessions.get(session_id)
        if (session is None or session.patient_id != target.patient_id
                or session.therapist_id != target.original_therapist_id
                or session.start_time.strftime("%H:%M") != target.original_time):
            raise ReplacementsReopenError("Pending target no longer joins to its immutable base session")

    try:
        fact_rows = build_daily_replacement_fact_rows(
            preflight.replacements, sessions=preflight.sessions,
        )
    except DailyReplacementFactError as exc:
        raise ReplacementsReopenError("Final active replacement facts are invalid") from exc
    fact_rows = tuple(sorted(fact_rows, key=lambda item: item.replacement_id))
    replacement_by_target = {row.target_session_id: row for row in fact_rows}
    if len(replacement_by_target) != len(fact_rows):
        raise ReplacementsReopenError("Duplicate active replacement target")
    for row in fact_rows:
        state = states[row.target_session_id]
        if (state.status != DailySessionStatus.REPLACED
                or state.replacement_id != row.replacement_id
                or state.effective_therapist_id != row.replacement_therapist_id
                or state.effective_time != row.replacement_time
                or state.effective_provider_kind != row.provider_kind):
            raise ReplacementsReopenError("Active replacement fact differs from effective daily state")

    suspensions: list[DailySessionCancellation] = []
    seen_cancellations: set[str] = set()
    for cancellation in preflight.daily_input.cancellations:
        if type(cancellation) is not DailySessionCancellation:
            raise ReplacementsReopenError("Invalid DAILY_INPUT cancellation type")
        if cancellation.cancellation_date != preflight.service_date:
            raise ReplacementsReopenError("Cancellation date differs from final service date")
        if cancellation.cancellation_id in seen_cancellations:
            raise ReplacementsReopenError("Duplicate DAILY_INPUT cancellation identity")
        seen_cancellations.add(cancellation.cancellation_id)
        if cancellation.kind == SessionCancellationKind.TREATMENT_SUSPENDED:
            session = sessions.get(cancellation.target_session_id)
            if session is None:
                raise ReplacementsReopenError("Treatment suspension targets an unknown session")
            state = states[cancellation.target_session_id]
            if (state.status != DailySessionStatus.CANCELLED
                    or state.cancellation_id != cancellation.cancellation_id
                    or state.cancellation_kind != SessionCancellationKind.TREATMENT_SUSPENDED):
                raise ReplacementsReopenError("Treatment suspension differs from effective daily state")
            suspensions.append(cancellation)

    suspended_by_target = {item.target_session_id: item for item in suspensions}
    if len(suspended_by_target) != len(suspensions):
        raise ReplacementsReopenError("Multiple treatment suspensions target one session")
    if set(replacement_by_target) & set(suspended_by_target):
        raise ReplacementsReopenError("A session cannot be both replaced and suspended")

    pending_replaced = set(pending) & set(replacement_by_target)
    pending_suspended = set(pending) & set(suspended_by_target)
    if pending_replaced & pending_suspended or pending_replaced | pending_suspended != set(pending):
        raise ReplacementsReopenError("Every pending target must be replaced or explicitly suspended exactly once")

    replacement_payloads = tuple(_replacement_payload(item) for item in fact_rows)
    suspension_payloads = tuple(
        ReopenedSuspension(
            cancellation_id=item.cancellation_id,
            target_session_id=item.target_session_id,
            patient_id=sessions[item.target_session_id].patient_id,
            original_therapist_id=sessions[item.target_session_id].therapist_id,
            cancellation_date=item.cancellation_date.isoformat(),
            cancellation_kind=item.kind.value,
        )
        for item in sorted(suspensions, key=lambda item: item.target_session_id)
    )
    unsigned = {
        "schema": SCHEMA,
        "operation_id": operation.operation_id,
        "operation_source_sha256": operation.source_sha256,
        "service_date": operation.service_date,
        "prior_inventory_sha256": pending_inventory.inventory_sha256,
        "prior_inventory_workbook_sha256": pending_inventory.evidence_workbook_sha256,
        "evidence_workbook_sha256": preflight.source_workbook_sha256,
        "pending_target_count": len(pending),
        "pending_replaced_count": len(pending_replaced),
        "pending_suspended_count": len(pending_suspended),
        "replacement_count": len(replacement_payloads),
        "suspension_count": len(suspension_payloads),
        "replacements": [asdict(item) for item in replacement_payloads],
        "suspensions": [asdict(item) for item in suspension_payloads],
    }
    digest = sha256(_canonical_json(unsigned).encode("ascii")).hexdigest()
    return ReplacementsReopenEvidence(
        **{key: unsigned[key] for key in (
            "schema", "operation_id", "operation_source_sha256", "service_date",
            "prior_inventory_sha256", "prior_inventory_workbook_sha256",
            "evidence_workbook_sha256", "pending_target_count",
            "pending_replaced_count", "pending_suspended_count",
            "replacement_count", "suspension_count",
        )},
        replacements=replacement_payloads,
        suspensions=suspension_payloads,
        evidence_sha256=digest,
    )


def validate_replacements_reopen_evidence(evidence: ReplacementsReopenEvidence) -> None:
    if type(evidence) is not ReplacementsReopenEvidence or evidence.schema != SCHEMA:
        raise ReplacementsReopenError("Exact supported replacement evidence required")
    for value in (
        evidence.operation_source_sha256, evidence.prior_inventory_sha256,
        evidence.prior_inventory_workbook_sha256, evidence.evidence_workbook_sha256,
        evidence.evidence_sha256,
    ):
        if type(value) is not str or not _HEX.fullmatch(value):
            raise ReplacementsReopenError("Replacement evidence digest is invalid")
    if (type(evidence.operation_id) is not str or not evidence.operation_id.strip()
            or type(evidence.service_date) is not str):
        raise ReplacementsReopenError("Replacement evidence identity is invalid")
    try:
        if date.fromisoformat(evidence.service_date).isoformat() != evidence.service_date:
            raise ValueError
    except ValueError as exc:
        raise ReplacementsReopenError("Replacement evidence date is invalid") from exc
    counts = (
        evidence.pending_target_count, evidence.pending_replaced_count,
        evidence.pending_suspended_count, evidence.replacement_count,
        evidence.suspension_count,
    )
    if (any(type(value) is not int or value < 0 for value in counts)
            or evidence.pending_target_count < 1
            or evidence.pending_replaced_count + evidence.pending_suspended_count
               != evidence.pending_target_count
            or type(evidence.replacements) is not tuple
            or type(evidence.suspensions) is not tuple
            or evidence.replacement_count != len(evidence.replacements)
            or evidence.suspension_count != len(evidence.suspensions)):
        raise ReplacementsReopenError("Replacement evidence counts are inconsistent")
    if any(type(item) is not ReopenedReplacement for item in evidence.replacements):
        raise ReplacementsReopenError("Replacement evidence row type is invalid")
    if any(type(item) is not ReopenedSuspension for item in evidence.suspensions):
        raise ReplacementsReopenError("Suspension evidence row type is invalid")
    pending_ids = {item.target_session_id for item in evidence.replacements}
    suspension_ids = {item.target_session_id for item in evidence.suspensions}
    if len(pending_ids) != len(evidence.replacements) or len(suspension_ids) != len(evidence.suspensions):
        raise ReplacementsReopenError("Duplicate replacement or suspension target")
    expected = sha256(_canonical_json(evidence.unsigned_dict()).encode("ascii")).hexdigest()
    if evidence.evidence_sha256 != expected:
        raise ReplacementsReopenError("Replacement evidence SHA-256 mismatch")


def write_replacements_reopen_artifact(
    evidence: ReplacementsReopenEvidence, *, evidence_root: str | Path,
    artifact_path: str | Path,
) -> ReplacementsReopenArtifact:
    validate_replacements_reopen_evidence(evidence)
    try:
        root = Path(evidence_root).resolve(strict=True)
        parent = Path(artifact_path).parent.resolve(strict=True)
        name = Path(artifact_path).name
        if not root.is_dir() or not name or name in {".", ".."}:
            raise ReplacementsReopenError("Invalid evidence root or artifact filename")
        locator = (parent.relative_to(root) / name).as_posix()
        if locator in {"", "."} or any(part in {"", ".", ".."} for part in Path(locator).parts):
            raise ReplacementsReopenError("Unsafe evidence artifact locator")
        output = parent / name
    except ReplacementsReopenError:
        raise
    except (OSError, ValueError) as exc:
        raise ReplacementsReopenError("Evidence artifact path escapes evidence root") from exc
    payload = evidence.to_json().encode("utf-8")
    digest = sha256(payload).hexdigest()
    created = False
    try:
        with output.open("xb") as handle:
            created = True
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
    except FileExistsError as exc:
        raise ReplacementsReopenError("Evidence artifact already exists") from exc
    except OSError as exc:
        if created:
            output.unlink(missing_ok=True)
        raise ReplacementsReopenError("Could not persist replacement evidence artifact") from exc
    return ReplacementsReopenArtifact(locator, digest, len(payload))


def read_replacements_reopen_artifact(
    *, evidence_root: str | Path, artifact_path: str | Path,
) -> tuple[ReplacementsReopenEvidence, ReplacementsReopenArtifact]:
    try:
        root = Path(evidence_root).resolve(strict=True)
        candidate = Path(artifact_path).resolve(strict=True)
        locator = candidate.relative_to(root).as_posix()
        if not root.is_dir() or not locator or locator == "." or not candidate.is_file():
            raise ValueError
        payload = candidate.read_bytes()
        data = json.loads(payload.decode("utf-8"))
    except (OSError, ValueError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReplacementsReopenError("Replacement evidence artifact is missing, unsafe or invalid") from exc
    fields = {
        "schema", "operation_id", "operation_source_sha256", "service_date",
        "prior_inventory_sha256", "prior_inventory_workbook_sha256",
        "evidence_workbook_sha256", "pending_target_count", "pending_replaced_count",
        "pending_suspended_count", "replacement_count", "suspension_count",
        "replacements", "suspensions", "evidence_sha256",
    }
    replacement_fields = {
        "replacement_id", "target_session_id", "patient_id", "original_therapist_id",
        "replacement_therapist_id", "replacement_date", "replacement_time", "provider_kind",
    }
    suspension_fields = {
        "cancellation_id", "target_session_id", "patient_id", "original_therapist_id",
        "cancellation_date", "cancellation_kind",
    }
    if (type(data) is not dict or set(data) != fields
            or type(data.get("replacements")) is not list
            or type(data.get("suspensions")) is not list):
        raise ReplacementsReopenError("Replacement evidence fields are missing or unexpected")
    if any(type(item) is not dict or set(item) != replacement_fields for item in data["replacements"]):
        raise ReplacementsReopenError("Replacement evidence row fields are invalid")
    if any(type(item) is not dict or set(item) != suspension_fields for item in data["suspensions"]):
        raise ReplacementsReopenError("Suspension evidence row fields are invalid")
    try:
        normalized = dict(data)
        normalized["replacements"] = tuple(ReopenedReplacement(**item) for item in data["replacements"])
        normalized["suspensions"] = tuple(ReopenedSuspension(**item) for item in data["suspensions"])
        evidence = ReplacementsReopenEvidence(**normalized)
    except TypeError as exc:
        raise ReplacementsReopenError("Replacement evidence field types are invalid") from exc
    validate_replacements_reopen_evidence(evidence)
    if evidence.to_json().encode("utf-8") != payload:
        raise ReplacementsReopenError("Replacement evidence artifact bytes are not canonical")
    return evidence, ReplacementsReopenArtifact(
        artifact_locator=locator, artifact_sha256=sha256(payload).hexdigest(),
        byte_count=len(payload),
    )
