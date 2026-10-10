from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import traceback

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.registration_transaction import (  # noqa: E402
    RegistrationTransactionError,
    run_registration_transaction,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Run one complete async registration transaction: close workbook, "
            "build verified preview, commit, and reopen."
        )
    )
    parser.add_argument("--request", type=Path, required=True)
    parser.add_argument("--response", type=Path, required=True)
    return parser


def _write_response(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _run_transaction_with_com(payload: dict[str, object]):
    """Run the detached transaction inside an initialized COM apartment on Windows."""

    if sys.platform != "win32":
        return run_registration_transaction(payload)

    try:
        import pythoncom  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RegistrationTransactionError(
            "pywin32/pythoncom is required for the Windows registration worker"
        ) from exc

    pythoncom.CoInitialize()
    try:
        return run_registration_transaction(payload)
    finally:
        pythoncom.CoUninitialize()


def main() -> int:
    args = _build_parser().parse_args()
    try:
        payload = json.loads(args.request.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, dict):
            raise RegistrationTransactionError("request JSON must be an object")
        report = _run_transaction_with_com(payload)
    except Exception as exc:  # Worker is detached; never let failures disappear silently.
        try:
            _write_response(
                args.response,
                {
                    "ok": False,
                    "error": str(exc),
                    "error_type": type(exc).__name__,
                    "traceback": traceback.format_exc(),
                },
            )
        except Exception:
            pass
        return 2

    _write_response(
        args.response,
        {
            "ok": True,
            "subject_key": report.subject_key,
            "display_name": report.display_name,
            "source_path": report.source_path,
            "preview_path": report.preview_path,
            "committed": report.committed,
            "reopened": report.reopened,
            "close_attempts": report.close_attempts,
            "timings_seconds": {
                "close": round(report.close_seconds, 3),
                "preview": round(report.preview_seconds, 3),
                "commit_reopen": round(report.commit_reopen_seconds, 3),
                "total": round(report.total_seconds, 3),
            },
        },
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
