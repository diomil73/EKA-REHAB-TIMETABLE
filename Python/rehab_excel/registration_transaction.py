from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import time
from typing import Callable, Mapping
from uuid import uuid4

from .authoritative_commit_worker import (
    AuthoritativeCommitWorkerError,
    AuthoritativeCommitWorkerReport,
    _default_close_workbook,
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


SleepService = Callable[[float], None]
CloseService = Callable[[Path], None]


def _close_with_retry(
    source: Path,
    *,
    timeout_seconds: float,
    poll_seconds: float,
    close_service: CloseService,
    sleep_service: SleepService,
) -> int:
    deadline = time.monotonic() + timeout_seconds
    attempts = 0

    while True:
        attempts += 1
        try:
            close_service(source)
            return attempts
        except Exception as exc:
            if time.monotonic() >= deadline:
                raise RegistrationTransactionError(
                    f"Could not close authoritative workbook before registration: {exc}"
                ) from exc
            sleep_service(poll_seconds)


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
    *,
    close_delay_seconds: float = 1.0,
    close_timeout_seconds: float = 10.0,
    close_poll_seconds: float = 0.5,
    sleep_service: SleepService = time.sleep,
    close_service: CloseService = _default_close_workbook,
) -> RegistrationTransactionReport:
    """Own one registration transaction after the UserForm event returns.

    The worker asks the workbook to close itself through its own ExitApplication
    macro, rather than issuing an external COM Workbook.Close. After shutdown it
    builds the isolated verified preview, commits it, and reopens the workbook.
    """

    source_text = str(payload.get("source_path", "")).strip()
    if not source_text:
        raise RegistrationTransactionError("source_path is required")

    source = Path(source_text).resolve()
    if close_delay_seconds < 0:
        raise RegistrationTransactionError("close_delay_seconds must be non-negative")
    if close_timeout_seconds <= 0:
        raise RegistrationTransactionError("close_timeout_seconds must be positive")
    if close_poll_seconds <= 0:
        raise RegistrationTransactionError("close_poll_seconds must be positive")

    transaction_payload, transaction_dir = _transaction_payload(payload)

    if close_delay_seconds:
        sleep_service(close_delay_seconds)

    close_attempts = _close_with_retry(
        source,
        timeout_seconds=close_timeout_seconds,
        poll_seconds=close_poll_seconds,
        close_service=close_service,
        sleep_service=sleep_service,
    )

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
        close_attempts=close_attempts,
    )
