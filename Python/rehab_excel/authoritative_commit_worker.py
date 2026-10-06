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
CloseService = Callable[[Path], None]


def _default_open_workbook(path: Path) -> None:
    if os.name != "nt":
        raise AuthoritativeCommitWorkerError(
            "Automatic workbook reopen is only supported on Windows"
        )
    os.startfile(str(path))  # type: ignore[attr-defined]


def _default_close_workbook(path: Path) -> None:
    if os.name != "nt":
        raise AuthoritativeCommitWorkerError(
            "Automatic workbook close is only supported on Windows"
        )

    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise AuthoritativeCommitWorkerError(
            "Automatic workbook close requires pywin32"
        ) from exc

    pythoncom.CoInitialize()
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            # Excel may already have closed between worker startup and this check.
            return

        target = str(path.resolve()).casefold()
        for index in range(1, int(excel.Workbooks.Count) + 1):
            workbook = excel.Workbooks(index)
            try:
                full_name = str(workbook.FullName or "").casefold()
            except Exception:
                continue
            if full_name == target:
                workbook.Close(SaveChanges=False)
                return
    finally:
        pythoncom.CoUninitialize()


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
    post_reopen_target: str | None = None,
    close_open_workbook: bool = False,
    close_delay_seconds: float = 1.5,
    commit_service: CommitService = commit_verified_preview,
    sleep_service: SleepService = time.sleep,
    open_service: OpenService = _default_open_workbook,
    close_service: CloseService = _default_close_workbook,
) -> AuthoritativeCommitWorkerReport:
    if timeout_seconds <= 0:
        raise AuthoritativeCommitWorkerError("timeout_seconds must be positive")
    if poll_seconds <= 0:
        raise AuthoritativeCommitWorkerError("poll_seconds must be positive")
    if close_delay_seconds < 0:
        raise AuthoritativeCommitWorkerError("close_delay_seconds must be non-negative")

    source = Path(source_path).resolve()

    if close_open_workbook:
        if close_delay_seconds:
            sleep_service(close_delay_seconds)
        try:
            close_service(source)
        except Exception as exc:
            raise AuthoritativeCommitWorkerError(
                f"Could not close authoritative workbook before commit: {exc}"
            ) from exc

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
            if post_reopen_target:
                sidecar = source.with_name(source.name + ".eka_next_sheet")
                sidecar.write_text(post_reopen_target.strip(), encoding="utf-8")
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
