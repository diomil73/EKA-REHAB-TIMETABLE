from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
from typing import Callable, Mapping

from .authoritative_commit import (
    AuthoritativeCommitError,
    AuthoritativeCommitReport,
    commit_verified_preview,
)


class AuthoritativeCommitBridgeError(RuntimeError):
    """Stable UI-facing error for authoritative workbook promotion."""


CommitService = Callable[..., AuthoritativeCommitReport]


def run_authoritative_commit_bridge(
    payload: Mapping[str, object],
    *,
    commit_service: CommitService = commit_verified_preview,
) -> dict[str, object]:
    source_path = str(payload.get("source_path", "")).strip()
    preview_path = str(payload.get("preview_path", "")).strip()
    expected_source_sha256 = str(payload.get("expected_source_sha256", "")).strip()
    backup_path_raw = str(payload.get("backup_path", "")).strip()
    remove_preview = payload.get("remove_preview_after_success", True)

    if not source_path:
        raise AuthoritativeCommitBridgeError("source_path is required")
    if not preview_path:
        raise AuthoritativeCommitBridgeError("preview_path is required")
    if not expected_source_sha256:
        raise AuthoritativeCommitBridgeError("expected_source_sha256 is required")
    if not isinstance(remove_preview, bool):
        raise AuthoritativeCommitBridgeError(
            "remove_preview_after_success must be a boolean"
        )

    backup_path = Path(backup_path_raw) if backup_path_raw else None

    try:
        report = commit_service(
            source_path,
            preview_path,
            expected_source_sha256=expected_source_sha256,
            backup_path=backup_path,
            remove_preview_after_success=remove_preview,
        )
    except (AuthoritativeCommitError, OSError, ValueError) as exc:
        raise AuthoritativeCommitBridgeError(str(exc)) from exc

    response = asdict(report)
    response["ok"] = True
    return response
