"""Non-authorizing structural codec for DAILY_INPUT transaction receipts.

This module validates receipt shape, canonical digest, operation/journal binding,
verifier registry labels and evidence artifact bytes. It does NOT authenticate
the receipt creator or verifier. SHA-256 is not a signature. No caller may use
this module alone to enable Apply or authorize a journal phase transition.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date, datetime, timezone
from hashlib import sha256
import json
from pathlib import Path, PurePosixPath
import re

from .daily_input_operation import DailyInputOperation
from .daily_input_phase_evidence import EvidenceKind

RECEIPT_SCHEMA = "diomil-transaction-receipt-v1"
_HEX = re.compile(r"^[0-9a-f]{64}$")
_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$")
_FIELDS = frozenset({
    "schema", "receipt_id", "operation_id", "source_workbook_sha256",
    "service_date", "journal_revision", "journal_phase", "evidence_kind",
    "evidence_artifact", "evidence_sha256", "verifier_id", "verifier_version",
    "verification_method", "verification_result", "verified_at_utc",
    "previous_receipt_sha256", "receipt_sha256",
})


class TransactionReceiptError(ValueError):
    """Raised when a receipt is malformed, unbound or inconsistent."""


def _canonical_json(value: dict) -> str:
    return json.dumps(value, ensure_ascii=True, allow_nan=False, sort_keys=True,
                      separators=(",", ":"))


def _require_text(name: str, value: object, *, identifier: bool = False) -> str:
    if type(value) is not str or not value.strip():
        raise TransactionReceiptError(f"{name} must be a non-empty string")
    if identifier and not _ID.fullmatch(value):
        raise TransactionReceiptError(f"{name} has invalid identifier syntax")
    return value


def _require_hash(name: str, value: object, *, nullable: bool = False) -> str | None:
    if nullable and value is None:
        return None
    if type(value) is not str or not _HEX.fullmatch(value):
        raise TransactionReceiptError(f"{name} must be a lowercase SHA-256 digest")
    return value


def _validate_artifact_locator(value: object) -> str:
    text = _require_text("evidence_artifact", value)
    if "\\" in text or text.startswith("/") or ":" in text:
        raise TransactionReceiptError("evidence_artifact must be a relative POSIX locator")
    path = PurePosixPath(text)
    if path.is_absolute() or any(part in {"", ".", ".."} for part in path.parts):
        raise TransactionReceiptError("evidence_artifact contains unsafe path segments")
    if path.as_posix() != text:
        raise TransactionReceiptError("evidence_artifact is not canonical")
    return text


def _validate_timestamp(value: object) -> str:
    text = _require_text("verified_at_utc", value)
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as exc:
        raise TransactionReceiptError("verified_at_utc must be RFC3339") from exc
    if parsed.tzinfo is None or parsed.utcoffset() != timezone.utc.utcoffset(parsed):
        raise TransactionReceiptError("verified_at_utc must be expressed in UTC")
    if not text.endswith("Z") and not text.endswith("+00:00"):
        raise TransactionReceiptError("verified_at_utc must use Z or +00:00")
    return text


@dataclass(frozen=True)
class DailyInputTransactionReceipt:
    schema: str
    receipt_id: str
    operation_id: str
    source_workbook_sha256: str
    service_date: str
    journal_revision: int
    journal_phase: str
    evidence_kind: str
    evidence_artifact: str
    evidence_sha256: str
    verifier_id: str
    verifier_version: str
    verification_method: str
    verification_result: str
    verified_at_utc: str
    previous_receipt_sha256: str | None
    receipt_sha256: str

    def __post_init__(self) -> None:
        if self.schema != RECEIPT_SCHEMA:
            raise TransactionReceiptError("Unsupported receipt schema")
        _require_text("receipt_id", self.receipt_id, identifier=True)
        _require_text("operation_id", self.operation_id, identifier=True)
        _require_hash("source_workbook_sha256", self.source_workbook_sha256)
        _require_hash("evidence_sha256", self.evidence_sha256)
        _require_hash("receipt_sha256", self.receipt_sha256)
        _require_hash("previous_receipt_sha256", self.previous_receipt_sha256, nullable=True)
        _require_text("service_date", self.service_date)
        try:
            if date.fromisoformat(self.service_date).isoformat() != self.service_date:
                raise ValueError
        except ValueError as exc:
            raise TransactionReceiptError("service_date must be YYYY-MM-DD") from exc
        if type(self.journal_revision) is not int or self.journal_revision < 0:
            raise TransactionReceiptError("journal_revision must be a non-negative integer")
        _require_text("journal_phase", self.journal_phase)
        try:
            EvidenceKind(self.evidence_kind)
        except ValueError as exc:
            raise TransactionReceiptError("Unknown evidence_kind") from exc
        _validate_artifact_locator(self.evidence_artifact)
        _require_text("verifier_id", self.verifier_id, identifier=True)
        _require_text("verifier_version", self.verifier_version, identifier=True)
        _require_text("verification_method", self.verification_method, identifier=True)
        if self.verification_result != "VERIFIED":
            raise TransactionReceiptError("Only VERIFIED evidence can have a success receipt")
        _validate_timestamp(self.verified_at_utc)
        if self.journal_revision == 0:
            if self.journal_phase != "PREPARED" or self.previous_receipt_sha256 is not None:
                raise TransactionReceiptError("Revision zero must be PREPARED with no prior receipt")
        elif self.previous_receipt_sha256 is None:
            raise TransactionReceiptError("Non-initial receipt requires previous_receipt_sha256")

    def unsigned_dict(self) -> dict:
        value = asdict(self)
        value.pop("receipt_sha256")
        return value

    def to_dict(self) -> dict:
        return asdict(self)

    def to_json(self) -> str:
        return _canonical_json(self.to_dict())


def create_transaction_receipt(**fields) -> DailyInputTransactionReceipt:
    """Create a structurally valid receipt; this does not authenticate its origin."""
    unexpected = set(fields) - (_FIELDS - {"receipt_sha256"})
    missing = (_FIELDS - {"receipt_sha256"}) - set(fields)
    if unexpected or missing:
        raise TransactionReceiptError("Receipt fields missing or unexpected")
    unsigned = {"schema": RECEIPT_SCHEMA, **fields}
    if "schema" in fields and fields["schema"] != RECEIPT_SCHEMA:
        raise TransactionReceiptError("Unsupported receipt schema")
    digest = sha256(_canonical_json(unsigned).encode("ascii")).hexdigest()
    return DailyInputTransactionReceipt(**unsigned, receipt_sha256=digest)


def decode_transaction_receipt(value: str | bytes | dict) -> DailyInputTransactionReceipt:
    """Decode strict canonical JSON and verify its self-digest (not authenticity)."""
    try:
        if isinstance(value, bytes):
            value = value.decode("utf-8")
        data = json.loads(value) if isinstance(value, str) else value
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise TransactionReceiptError("Receipt JSON is invalid") from exc
    if type(data) is not dict or set(data) != _FIELDS:
        raise TransactionReceiptError("Receipt schema fields are missing or unexpected")
    receipt = DailyInputTransactionReceipt(**data)
    expected = sha256(_canonical_json(receipt.unsigned_dict()).encode("ascii")).hexdigest()
    if receipt.receipt_sha256 != expected:
        raise TransactionReceiptError("Receipt SHA-256 mismatch")
    if isinstance(value, str) and _canonical_json(receipt.to_dict()) != value:
        raise TransactionReceiptError("Receipt JSON is not canonical")
    return receipt


def validate_receipt_binding(
    receipt: DailyInputTransactionReceipt,
    operation: DailyInputOperation,
    *,
    required_evidence: tuple[EvidenceKind, ...],
    expected_previous_receipt_sha256: str | None,
    verifier_registry: dict[str, str],
) -> None:
    """Check structural binding to a decoded operation and pinned verifier label.

    This is not a trust/authenticity decision. The registry mapping must itself
    come from a separately trusted application build manifest.
    """
    if type(receipt) is not DailyInputTransactionReceipt or type(operation) is not DailyInputOperation:
        raise TransactionReceiptError("Invalid receipt or operation type")
    if (receipt.operation_id, receipt.source_workbook_sha256, receipt.service_date,
            receipt.journal_revision, receipt.journal_phase) != (
            operation.operation_id, operation.source_sha256, operation.service_date,
            operation.revision, operation.phase.value):
        raise TransactionReceiptError("Receipt does not match journal operation identity/revision")
    if receipt.evidence_kind not in {item.value for item in required_evidence}:
        raise TransactionReceiptError("Receipt evidence kind does not satisfy phase contract")
    if receipt.previous_receipt_sha256 != expected_previous_receipt_sha256:
        raise TransactionReceiptError("Receipt chain predecessor mismatch")
    if verifier_registry.get(receipt.verifier_id) != receipt.verifier_version:
        raise TransactionReceiptError("Verifier identity/version is not pinned in registry")


def verify_receipt_artifact_bytes(
    receipt: DailyInputTransactionReceipt,
    evidence_root: str | Path,
) -> str:
    """Verify the exact evidence artifact bytes under root; returns its SHA-256.

    This verifies bytes and containment, not the authenticity of the verifier.
    """
    root = Path(evidence_root).resolve(strict=True)
    candidate = root.joinpath(*PurePosixPath(receipt.evidence_artifact).parts)
    try:
        resolved = candidate.resolve(strict=True)
        resolved.relative_to(root)
    except (OSError, ValueError) as exc:
        raise TransactionReceiptError("Evidence artifact missing or escapes evidence root") from exc
    if not resolved.is_file():
        raise TransactionReceiptError("Evidence artifact is not a regular file")
    digest = sha256(resolved.read_bytes()).hexdigest()
    if digest != receipt.evidence_sha256:
        raise TransactionReceiptError("Evidence artifact SHA-256 mismatch")
    return digest
