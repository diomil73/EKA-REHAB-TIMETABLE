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


class _SharedWorkbook:
    """Delegate one COM workbook while deferring physical close until preview end."""

    def __init__(self, workbook) -> None:
        object.__setattr__(self, "_workbook", workbook)

    def __getattr__(self, name: str):
        return getattr(object.__getattribute__(self, "_workbook"), name)

    def __setattr__(self, name: str, value) -> None:
        setattr(object.__getattribute__(self, "_workbook"), name, value)

    def Close(self, *args, **kwargs) -> None:
        # Existing preview stages close the workbook in their own finally blocks.
        # Suppress those closes so the next stage can reuse the same open file.
        return None

    def _force_close(self) -> None:
        workbook = object.__getattribute__(self, "_workbook")
        workbook.Close(SaveChanges=False)


class _SharedWorkbooks:
    """Cache Workbooks.Open by resolved path for one preview transaction."""

    def __init__(self, workbooks) -> None:
        object.__setattr__(self, "_workbooks", workbooks)
        object.__setattr__(self, "_opened", {})

    def __getattr__(self, name: str):
        return getattr(object.__getattribute__(self, "_workbooks"), name)

    @staticmethod
    def _key(path: object) -> str:
        return str(Path(str(path)).resolve()).casefold()

    def Open(self, path, *args, **kwargs):
        opened = object.__getattribute__(self, "_opened")
        key = self._key(path)
        if key in opened:
            return opened[key]

        workbooks = object.__getattribute__(self, "_workbooks")
        workbook = workbooks.Open(path, *args, **kwargs)
        proxy = _SharedWorkbook(workbook)
        opened[key] = proxy
        return proxy

    def _close_all(self) -> None:
        opened = object.__getattribute__(self, "_opened")
        for workbook in reversed(tuple(opened.values())):
            try:
                workbook._force_close()
            except Exception:
                pass
        opened.clear()


class _SharedExcelApplication:
    """Proxy one hidden Excel instance and its open workbooks across preview stages."""

    def __init__(self, application) -> None:
        object.__setattr__(self, "_application", application)
        object.__setattr__(self, "_workbooks", _SharedWorkbooks(application.Workbooks))

    def __getattr__(self, name: str):
        if name == "Workbooks":
            return object.__getattribute__(self, "_workbooks")
        return getattr(object.__getattribute__(self, "_application"), name)

    def __setattr__(self, name: str, value) -> None:
        setattr(object.__getattribute__(self, "_application"), name, value)

    def Quit(self) -> None:
        # Stage-level cleanup must not terminate the shared Excel process.
        return None

    def _close_all_workbooks(self) -> None:
        object.__getattribute__(self, "_workbooks")._close_all()


@dataclass(frozen=True)
class RegistrationTransactionReport:
    subject_key: str
    display_name: str
    source_path: str
    preview_path: str
    committed: bool
    reopened: bool
    close_attempts: int
    close_seconds: float
    preview_seconds: float
    commit_reopen_seconds: float
    total_seconds: float


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
    """Build the preview in one COM apartment, Excel process, and open workbook.

    Patient preview creation touches PATIENTS, PATIENT_PLANNER, MASTER and visual
    styling in sequence. Each stage still saves normally, so file-based safety
    checks and readers observe committed stage output, but repeated Workbooks.Open
    and workbook Close calls are collapsed into one physical open/close cycle.
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
    shared_excel = None
    original_dispatch_ex = None
    try:
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

        win32com.client.DispatchEx = shared_dispatch_ex
        return run_registration_bridge(payload)
    finally:
        if original_dispatch_ex is not None:
            try:
                win32com.client.DispatchEx = original_dispatch_ex
            except Exception:
                pass
        if shared_excel is not None:
            try:
                shared_excel._close_all_workbooks()
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
    """Close the saved workbook, build a verified preview, commit, and reopen."""

    total_started = time.monotonic()

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

    close_started = time.monotonic()
    close_attempts = _close_with_retry(
        source,
        timeout_seconds=close_timeout_seconds,
        poll_seconds=close_poll_seconds,
        close_service=close_service,
        sleep_service=sleep_service,
    )
    close_seconds = time.monotonic() - close_started

    try:
        preview_started = time.monotonic()
        preview = _run_bridge_with_fresh_com(transaction_payload)
        preview_seconds = time.monotonic() - preview_started

        commit_started = time.monotonic()
        commit: AuthoritativeCommitWorkerReport = commit_when_unlocked(
            source,
            str(preview["output_path"]),
            expected_source_sha256=str(preview["source_sha256_before"]),
            remove_preview_after_success=True,
            timeout_seconds=30.0,
            poll_seconds=0.5,
            reopen=True,
        )
        commit_reopen_seconds = time.monotonic() - commit_started
    except (RegistrationBridgeError, AuthoritativeCommitWorkerError, KeyError) as exc:
        raise RegistrationTransactionError(str(exc)) from exc

    if commit.commit.preview_removed:
        try:
            transaction_dir.rmdir()
            transaction_dir.parent.rmdir()
        except OSError:
            pass

    total_seconds = time.monotonic() - total_started
    return RegistrationTransactionReport(
        subject_key=str(preview.get("subject_key", "")),
        display_name=str(preview.get("display_name", "")),
        source_path=str(source),
        preview_path=str(preview.get("output_path", "")),
        committed=bool(commit.commit.committed),
        reopened=bool(commit.reopened),
        close_attempts=close_attempts,
        close_seconds=close_seconds,
        preview_seconds=preview_seconds,
        commit_reopen_seconds=commit_reopen_seconds,
        total_seconds=total_seconds,
    )
