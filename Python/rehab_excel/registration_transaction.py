from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
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


_ORIGINAL_RUN_REGISTRATION_BRIDGE = run_registration_bridge


class RegistrationTransactionError(RuntimeError):
    """Raised when the async registration transaction cannot complete safely."""


class _SharedExcelApplication:
    """Proxy one hidden Excel instance across all preview stages.

    Existing preview writers own the Excel instance they create and therefore call
    ``Quit`` in their cleanup blocks. During a registration transaction we want
    those writers to reuse one process instead. This proxy delegates everything
    except ``Quit`` so each stage can still close its workbook normally while the
    shared Excel process stays alive until the complete preview pipeline ends.
    """

    def __init__(self, application) -> None:
        object.__setattr__(self, "_application", application)

    def __getattr__(self, name: str):
        return getattr(object.__getattribute__(self, "_application"), name)

    def __setattr__(self, name: str, value) -> None:
        setattr(object.__getattribute__(self, "_application"), name, value)

    def Quit(self) -> None:
        return None


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


def _run_bridge_with_fresh_com(payload: Mapping[str, object]):
    """Build the preview in one fresh COM apartment and one Excel process.

    Patient preview creation touches PATIENTS, PATIENT_PLANNER, MASTER and the
    visual styling pass. Historically each stage launched and terminated its own
    hidden Excel process, which dominated the transaction time on the real
    workbook. Inside the detached worker we can safely reuse one isolated Excel
    process while preserving the existing stage-level workbook close/save logic.
    """

    if os.name != "nt":
        return run_registration_bridge(payload)

    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RegistrationTransactionError(
            "pywin32/pythoncom is required for Windows registration preview creation"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    original_dispatch_ex = None
    try:
        # Unit tests replace the bridge with a fake. Keep those tests independent
        # of a local Excel installation and exercise only the transaction order.
        if run_registration_bridge is not _ORIGINAL_RUN_REGISTRATION_BRIDGE:
            return run_registration_bridge(payload)

        original_dispatch_ex = win32com.client.DispatchEx
        excel = original_dispatch_ex("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        shared_excel = _SharedExcelApplication(excel)

        def shared_dispatch_ex(prog_id: str):
            if str(prog_id).strip().casefold() == "excel.application":
                return shared_excel
            return original_dispatch_ex(prog_id)

        # All preview modules import the same win32com.client module object. By
        # replacing DispatchEx only for the duration of this detached worker call,
        # their existing cleanup code becomes compatible with one shared process.
        win32com.client.DispatchEx = shared_dispatch_ex
        return run_registration_bridge(payload)
    finally:
        if original_dispatch_ex is not None:
            try:
                win32com.client.DispatchEx = original_dispatch_ex
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def run_registration_transaction(
    payload: Mapping[str, object],
    *,
    close_delay_seconds: float = 1.0,
    close_timeout_seconds: float = 10.0,
    close_poll_seconds: float = 0.5,
    sleep_service: SleepService = time.sleep,
    close_service: CloseService = _default_close_workbook,
) -> RegistrationTransactionReport:
    """Close the saved workbook, build a verified preview, commit, and reopen.

    The authoritative workbook must be closed before preview creation so Excel
    cannot rewrite the source after the preview captures its safety hash. This
    keeps the CAS check meaningful and prevents false commit refusals caused by
    Excel saving during shutdown.
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
        preview = _run_bridge_with_fresh_com(transaction_payload)
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
