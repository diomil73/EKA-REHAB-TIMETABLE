from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
import time
from typing import Callable

from .authoritative_commit import (
    AuthoritativeCommitError,
    AuthoritativeCommitReport,
    commit_verified_preview,
)


class AuthoritativeCommitWorkerError(RuntimeError):
    """Raised when a deferred authoritative commit cannot complete."""


@dataclass(frozen=True)
class AuthoritativeCommitWorkerReport:
    commit: AuthoritativeCommitReport
    attempts: int
    reopened: bool


CommitService = Callable[..., AuthoritativeCommitReport]
SleepService = Callable[[float], None]
OpenService = Callable[[Path], None]


def _default_open_workbook(path: Path) -> None:
    if os.name != "nt":
        raise AuthoritativeCommitWorkerError(
            "Automatic workbook reopen is only supported on Windows"
        )
    os.startfile(str(path))  # type: ignore[attr-defined]


def commit_when_unlocked(
    source_path: str | Path,
    preview_path: str | Path,
    *,
    expected_source_sha256: str,
    backup_path: str | Path | None = None,
    remove_preview_after_success: bool = True,
    timeout_seconds: float = 30.0,
    poll_seconds: float = 0.5,
    reopen: bool = False,
    commit_service: CommitService = commit_verified_preview,
    sleep_service: SleepService = time.sleep,
    open_service: OpenService = _default_open_workbook,
) -> AuthoritativeCommitWorkerReport:
    if timeout_seconds <= 0:
        raise AuthoritativeCommitWorkerError("timeout_seconds must be positive")
    if poll_seconds <= 0:
        raise AuthoritativeCommitWorkerError("poll_seconds must be positive")

    source = Path(source_path).resolve()
    deadline = time.monotonic() + timeout_seconds
    attempts = 0

    while True:
        attempts += 1
        try:
            report = commit_service(
                source,
                preview_path,
                expected_source_sha256=expected_source_sha256,
                backup_path=backup_path,
                remove_preview_after_success=remove_preview_after_success,
            )
            break
        except AuthoritativeCommitError as exc:
            if "open or locked by Excel" not in str(exc):
                raise AuthoritativeCommitWorkerError(str(exc)) from exc
            if time.monotonic() >= deadline:
                raise AuthoritativeCommitWorkerError(
                    "Timed out waiting for Excel to release the authoritative workbook"
                ) from exc
            sleep_service(poll_seconds)

    reopened = False
    if reopen:
        try:
            open_service(source)
            reopened = True
        except Exception as exc:
            raise AuthoritativeCommitWorkerError(
                f"Commit succeeded but workbook could not be reopened: {exc}"
            ) from exc

    return AuthoritativeCommitWorkerReport(
        commit=report,
        attempts=attempts,
        reopened=reopened,
    )
