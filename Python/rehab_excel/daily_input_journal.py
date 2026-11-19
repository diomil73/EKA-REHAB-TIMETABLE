"""Append-only, checksummed DAILY_INPUT operation journal codec.

Pure storage contract only. Workbook sheet creation, protection, append/save and
native reopen/readback belong to the Excel/VBA adapter and are not implemented
here. A digest detects accidental corruption; it is not an authenticity proof.
"""
from __future__ import annotations

from dataclasses import asdict
from hashlib import sha256
import json

from .daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationError,
    DailyInputOperationPhase,
    advance_daily_input_operation,
)
from .daily_input_day_switch import DailyInputDaySwitchError

JOURNAL_SHEET = "DAILY_INPUT_JOURNAL"
JOURNAL_HEADERS = ("OperationID", "Revision", "Phase", "PayloadSHA256", "Payload")
_PREFIX = "diomil-journal-v1:"


class DailyInputJournalError(DailyInputDaySwitchError):
    """Raised when append-only operation journal evidence is invalid."""


def _payload(operation: DailyInputOperation) -> str:
    value = {
        "operation_id": operation.operation_id,
        "source_sha256": operation.source_sha256,
        "service_date": operation.service_date,
        "phase": operation.phase.value,
        "revision": operation.revision,
        "failure_code": operation.failure_code,
        "failure_detail": operation.failure_detail,
        "operator_resolution": operation.operator_resolution,
        "rollback_verified": operation.rollback_verified,
    }
    return json.dumps(value, ensure_ascii=True, allow_nan=False,
                      sort_keys=True, separators=(",", ":"))


def _row(operation: DailyInputOperation) -> tuple:
    payload = _payload(operation)
    digest = sha256((_PREFIX + payload).encode("ascii")).hexdigest()
    return (operation.operation_id, operation.revision, operation.phase.value, digest, _PREFIX + payload)


def _decode_row(row) -> DailyInputOperation:
    if type(row) not in (tuple, list) or len(row) != 5:
        raise DailyInputJournalError("Invalid DAILY_INPUT journal row shape")
    operation_id, revision, phase, digest, stored = row
    if (type(operation_id) is not str or not operation_id.strip()
            or type(revision) is not int or revision < 0
            or type(phase) is not str or type(digest) is not str
            or len(digest) != 64
            or any(char not in "0123456789abcdef" for char in digest)
            or type(stored) is not str or not stored.startswith(_PREFIX)):
        raise DailyInputJournalError("Invalid DAILY_INPUT journal row fields")
    payload = stored[len(_PREFIX):]
    expected = sha256((_PREFIX + payload).encode("ascii")).hexdigest()
    if digest != expected:
        raise DailyInputJournalError("DAILY_INPUT journal checksum mismatch")
    try:
        value = json.loads(payload)
        if type(value) is not dict:
            raise ValueError("payload must be an object")
        expected_keys = {
            "operation_id", "source_sha256", "service_date", "phase", "revision",
            "failure_code", "failure_detail", "operator_resolution", "rollback_verified",
        }
        if set(value) != expected_keys:
            raise ValueError("unexpected journal payload schema")
        operation = DailyInputOperation(
            operation_id=value["operation_id"],
            source_sha256=value["source_sha256"],
            service_date=value["service_date"],
            phase=DailyInputOperationPhase(value["phase"]),
            revision=value["revision"],
            failure_code=value["failure_code"],
            failure_detail=value["failure_detail"],
            operator_resolution=value["operator_resolution"],
            rollback_verified=value["rollback_verified"],
        )
    except (ValueError, TypeError, KeyError, DailyInputOperationError) as exc:
        raise DailyInputJournalError("Invalid DAILY_INPUT journal payload") from exc
    if (operation.operation_id, operation.revision, operation.phase.value) != (
        operation_id, revision, phase
    ):
        raise DailyInputJournalError("Journal row/payload metadata mismatch")
    if _payload(operation) != payload:
        raise DailyInputJournalError("Journal payload is not canonical")
    return operation


def decode_operation_journal(rows) -> tuple[DailyInputOperation, ...]:
    """Validate every immutable revision and reconstruct its legal transition chain."""
    try:
        records = tuple(_decode_row(row) for row in rows)
        if not records:
            return ()
        first = records[0]
        if first.phase != DailyInputOperationPhase.PREPARED or first.revision != 0:
            raise DailyInputJournalError("Journal must begin with PREPARED revision 0")
        if any(record.operation_id != first.operation_id for record in records):
            raise DailyInputJournalError("Mixed operation IDs in one journal history")
        if any(record.source_sha256 != first.source_sha256
               or record.service_date != first.service_date for record in records):
            raise DailyInputJournalError("Operation identity changed inside journal history")
        current = first
        for candidate in records[1:]:
            if candidate.revision != current.revision + 1:
                raise DailyInputJournalError("Missing, duplicated, or reordered journal revision")
            kwargs = {}
            if candidate.phase == DailyInputOperationPhase.FAILED:
                kwargs = {"failure_code": candidate.failure_code,
                          "failure_detail": candidate.failure_detail}
            elif candidate.phase == DailyInputOperationPhase.ROLLED_BACK:
                kwargs = {"rollback_verified": candidate.rollback_verified}
            elif (current.phase == DailyInputOperationPhase.PENDING_OPERATOR_DECISION
                  and candidate.phase == DailyInputOperationPhase.REPLACEMENTS_RESOLVED):
                kwargs = {"operator_resolution": candidate.operator_resolution}
            reconstructed = advance_daily_input_operation(current, candidate.phase, **kwargs)
            if reconstructed != candidate:
                raise DailyInputJournalError("Journal transition payload differs from validated state")
            current = candidate
        return records
    except DailyInputOperationError as exc:
        raise DailyInputJournalError("Illegal DAILY_INPUT journal transition") from exc


def append_operation_journal(rows, operation: DailyInputOperation) -> tuple[tuple, ...]:
    """Return rows with one new revision; exact replay is idempotent, edits are rejected."""
    existing_rows = tuple(tuple(row) for row in rows)
    history = decode_operation_journal(existing_rows)
    if not history:
        if operation.phase != DailyInputOperationPhase.PREPARED or operation.revision != 0:
            raise DailyInputJournalError("First journal entry must be PREPARED revision 0")
        return (_row(operation),)
    current = history[-1]
    if operation == current:
        return existing_rows
    if operation.operation_id != current.operation_id:
        raise DailyInputJournalError("Cannot append a different operation to this journal")
    candidate_rows = existing_rows + (_row(operation),)
    decoded = decode_operation_journal(candidate_rows)
    if decoded[-1] != operation:
        raise DailyInputJournalError("Appended operation state did not read back exactly")
    return candidate_rows


def verify_operation_journal_readback(actual_rows, expected_rows) -> tuple[DailyInputOperation, ...]:
    """Verify full append-only history, not only the most recent journal row."""
    actual = tuple(tuple(row) for row in actual_rows)
    expected = tuple(tuple(row) for row in expected_rows)
    if actual != expected:
        raise DailyInputJournalError("DAILY_INPUT journal readback differs from expected history")
    return decode_operation_journal(actual)
