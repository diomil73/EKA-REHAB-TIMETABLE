from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import time

from tools.apply_day_amendment import DayAmendmentApplyError, apply_day_amendment


class DayAmendmentWorkerError(RuntimeError):
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
    attempts = 0
    last_error: Exception | None = None

    while time.monotonic() <= deadline:
        attempts += 1
        try:
            result = apply_day_amendment(workbook, payload)
            if reopen:
                os.startfile(str(workbook))  # type: ignore[attr-defined]
            return {**result, "attempts": attempts, "reopened": reopen}
        except (DayAmendmentApplyError, OSError, RuntimeError) as exc:
            last_error = exc
            if "Workbook changed after amendment preparation" in str(exc):
                raise DayAmendmentWorkerError(str(exc)) from exc
            time.sleep(poll_seconds)

    raise DayAmendmentWorkerError(
        f"Workbook was not released for amendment apply after {attempts} attempts: {last_error}"
    )


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
