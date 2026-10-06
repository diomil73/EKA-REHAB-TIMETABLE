from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Mapping
from uuid import uuid4

from .authoritative_commit_worker import (
    AuthoritativeCommitWorkerError,
    AuthoritativeCommitWorkerReport,
    commit_when_unlocked,
)
from .registration_bridge import RegistrationBridgeError, run_registration_bridge


class RegistrationTransactionError(RuntimeError):
    """Raised when the async registration transaction cannot complete safely."""


@dataclass(frozen=True)
class RegistrationTransactionReport:
    subject_key: str
    display_name: str
    source_path: str
    preview_path: str
    committed: bool
    reopened: bool
    close_attempts: int


def _transaction_payload(payload: Mapping[str, object]) -> tuple[dict[str, object], Path]:
    preview_root_text = str(payload.get("preview_dir", "")).strip()
    if not preview_root_text:
        raise RegistrationTransactionError("preview_dir is required")

    preview_root = Path(preview_root_text).resolve()
    transaction_dir = preview_root / ".eka_registration_transactions" / uuid4().hex
    transaction_payload = dict(payload)
    transaction_payload["preview_dir"] = str(transaction_dir)
    return transaction_payload, transaction_dir


def run_registration_transaction(
    payload: Mapping[str, object],
) -> RegistrationTransactionReport:
    """Build, commit, and reopen one registration transaction.

    Workbook shutdown is intentionally owned by VBA inside the workbook itself.
    The detached worker must never attempt an external COM Workbook.Close: Excel
    can refuse that call while a UserForm/event is unwinding. The VBA bridge
    schedules an internal Application.OnTime close after launching this worker.
    Preview creation may safely begin from the already-saved source, while the
    authoritative commit waits until Excel has released the source file.
    """

    source_text = str(payload.get("source_path", "")).strip()
    if not source_text:
        raise RegistrationTransactionError("source_path is required")

    source = Path(source_text).resolve()
    transaction_payload, transaction_dir = _transaction_payload(payload)

    try:
        preview = run_registration_bridge(transaction_payload)
        commit: AuthoritativeCommitWorkerReport = commit_when_unlocked(
            source,
            str(preview["output_path"]),
            expected_source_sha256=str(preview["source_sha256_before"]),
            remove_preview_after_success=True,
            timeout_seconds=30.0,
            poll_seconds=0.5,
            reopen=True,
        )
    except (RegistrationBridgeError, AuthoritativeCommitWorkerError, KeyError) as exc:
        raise RegistrationTransactionError(str(exc)) from exc

    if commit.commit.preview_removed:
        try:
            transaction_dir.rmdir()
            transaction_dir.parent.rmdir()
        except OSError:
            pass

    return RegistrationTransactionReport(
        subject_key=str(preview.get("subject_key", "")),
        display_name=str(preview.get("display_name", "")),
        source_path=str(source),
        preview_path=str(preview.get("output_path", "")),
        committed=bool(commit.commit.committed),
        reopened=bool(commit.reopened),
        close_attempts=0,
    )
