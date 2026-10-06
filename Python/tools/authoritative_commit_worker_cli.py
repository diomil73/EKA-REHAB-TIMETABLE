from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.authoritative_commit_worker import (  # noqa: E402
    AuthoritativeCommitWorkerError,
    commit_when_unlocked,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Wait for Excel to release an authoritative XLSM, commit a verified "
            "preview with backup, and optionally reopen the authoritative workbook."
        )
    )
    parser.add_argument("--request", type=Path, required=True)
    parser.add_argument("--response", type=Path, required=True)
    return parser


def _write_response(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> int:
    args = _build_parser().parse_args()

    try:
        payload = json.loads(args.request.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, dict):
            raise AuthoritativeCommitWorkerError("request JSON must be an object")

        source_path = str(payload.get("source_path", "")).strip()
        preview_path = str(payload.get("preview_path", "")).strip()
        expected_source_sha256 = str(payload.get("expected_source_sha256", "")).strip()
        backup_path_text = str(payload.get("backup_path", "")).strip()
        remove_preview = payload.get("remove_preview_after_success", True)
        reopen = payload.get("reopen", True)
        close_open_workbook = payload.get("close_open_workbook", False)
        timeout_seconds = float(payload.get("timeout_seconds", 30))
        poll_seconds = float(payload.get("poll_seconds", 0.5))
        close_delay_seconds = float(payload.get("close_delay_seconds", 1.5))
        post_reopen_target = str(payload.get("post_reopen_target", "")).strip()

        if not source_path:
            raise AuthoritativeCommitWorkerError("source_path is required")
        if not preview_path:
            raise AuthoritativeCommitWorkerError("preview_path is required")
        if not expected_source_sha256:
            raise AuthoritativeCommitWorkerError("expected_source_sha256 is required")
        if not isinstance(remove_preview, bool):
            raise AuthoritativeCommitWorkerError(
                "remove_preview_after_success must be a boolean"
            )
        if not isinstance(reopen, bool):
            raise AuthoritativeCommitWorkerError("reopen must be a boolean")
        if not isinstance(close_open_workbook, bool):
            raise AuthoritativeCommitWorkerError("close_open_workbook must be a boolean")

        report = commit_when_unlocked(
            source_path,
            preview_path,
            expected_source_sha256=expected_source_sha256,
            backup_path=backup_path_text or None,
            remove_preview_after_success=remove_preview,
            timeout_seconds=timeout_seconds,
            poll_seconds=poll_seconds,
            reopen=reopen,
            post_reopen_target=post_reopen_target or None,
            close_open_workbook=close_open_workbook,
            close_delay_seconds=close_delay_seconds,
        )
    except (OSError, ValueError, json.JSONDecodeError, AuthoritativeCommitWorkerError) as exc:
        _write_response(args.response, {"ok": False, "error": str(exc)})
        return 2

    _write_response(
        args.response,
        {
            "ok": True,
            "committed": report.commit.committed,
            "source_path": report.commit.source_path,
            "backup_path": report.commit.backup_path,
            "attempts": report.attempts,
            "reopened": report.reopened,
            "preview_removed": report.commit.preview_removed,
        },
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
