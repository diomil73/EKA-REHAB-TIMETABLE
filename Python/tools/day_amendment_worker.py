from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
import time

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.apply_day_amendment import DayAmendmentApplyError, apply_day_amendment  # noqa: E402


class DayAmendmentWorkerError(RuntimeError):
    pass


def _windows_error_code(exc: BaseException) -> int | None:
    code = getattr(exc, "winerror", None)
    if isinstance(code, int):
        return code
    if exc.args and isinstance(exc.args[0], int):
        return exc.args[0]
    return None


def _workbook_is_exclusively_available(path: Path) -> bool:
    """Return True only when Windows grants an exclusive read/write handle.

    This is intentionally a cheap file-lock probe. It must not start Excel while
    the interactive workbook is still closing, because a read-only Excel retry
    can itself prolong the handoff and obscure whether the original session has
    actually released the file.
    """

    if sys.platform != "win32":
        return True

    try:
        import win32con  # type: ignore[import-not-found]
        import win32file  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DayAmendmentWorkerError("pywin32 is required for workbook release probing") from exc

    handle = None
    try:
        handle = win32file.CreateFile(
            str(path),
            win32con.GENERIC_READ | win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None,
        )
        return True
    except Exception as exc:
        # pywin32 raises pywintypes.error for sharing/lock violations; that
        # exception is not guaranteed to inherit from OSError. These two Windows
        # errors simply mean that Excel still owns the workbook, so keep polling.
        if _windows_error_code(exc) in (32, 33):
            return False
        raise
    finally:
        if handle is not None:
            try:
                handle.Close()
            except Exception:
                pass


def apply_when_released(
    workbook_path: str | Path,
    payload_path: str | Path,
    *,
    timeout_seconds: float = 30.0,
    poll_seconds: float = 0.5,
    reopen: bool = True,
) -> dict[str, object]:
    workbook = Path(workbook_path).resolve()
    payload = Path(payload_path).resolve()
    deadline = time.monotonic() + timeout_seconds
    probes = 0

    while time.monotonic() <= deadline:
        probes += 1
        if _workbook_is_exclusively_available(workbook):
            break
        time.sleep(poll_seconds)
    else:
        raise DayAmendmentWorkerError(
            f"Workbook was not released for amendment apply after {probes} file-lock probes"
        )

    try:
        result = apply_day_amendment(workbook, payload)
    except DayAmendmentApplyError as exc:
        if "Workbook changed after amendment preparation" in str(exc):
            raise DayAmendmentWorkerError(str(exc)) from exc
        raise

    if reopen:
        os.startfile(str(workbook))  # type: ignore[attr-defined]
    return {**result, "attempts": probes, "reopened": reopen}


def _write_response(path: Path, data: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8-sig")


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply a prepared day amendment after Excel releases the workbook.")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--payload", required=True, type=Path)
    parser.add_argument("--response", required=True, type=Path)
    parser.add_argument("--timeout-seconds", type=float, default=30.0)
    parser.add_argument("--poll-seconds", type=float, default=0.5)
    parser.add_argument("--no-reopen", action="store_true")
    args = parser.parse_args()

    try:
        result = apply_when_released(
            args.workbook,
            args.payload,
            timeout_seconds=args.timeout_seconds,
            poll_seconds=args.poll_seconds,
            reopen=not args.no_reopen,
        )
    except Exception as exc:
        _write_response(args.response, {"ok": False, "error": str(exc)})
        return 2

    _write_response(args.response, result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
