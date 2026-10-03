from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import os
import shutil
import tempfile
from zipfile import ZipFile


class AuthoritativeCommitError(RuntimeError):
    """Raised when a verified preview cannot be promoted safely."""


@dataclass(frozen=True)
class AuthoritativeCommitReport:
    source_path: str
    preview_path: str
    backup_path: str
    source_sha256_before: str
    source_sha256_after: str
    preview_sha256: str
    committed: bool
    preview_removed: bool


def file_sha256(path: str | Path) -> str:
    target = Path(path)
    digest = sha256()
    with target.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def _default_backup_path(source: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}.backup_{stamp}{source.suffix}")


def commit_verified_preview(
    source_path: str | Path,
    preview_path: str | Path,
    *,
    expected_source_sha256: str,
    backup_path: str | Path | None = None,
    remove_preview_after_success: bool = True,
) -> AuthoritativeCommitReport:
    """Promote one already-verified XLSM preview to the authoritative workbook.

    The caller must supply the SHA-256 of the authoritative workbook captured
    before the preview was produced. Promotion is refused if the source changed
    in the meantime.

    The replacement is staged in the source directory and uses os.replace only
    after a backup has been created. If verification after replacement fails,
    the backup is restored.
    """

    source = Path(source_path).resolve()
    preview = Path(preview_path).resolve()

    if not source.exists():
        raise AuthoritativeCommitError(f"Authoritative workbook not found: {source}")
    if not preview.exists():
        raise AuthoritativeCommitError(f"Verified preview not found: {preview}")
    if source == preview:
        raise AuthoritativeCommitError("Preview must be different from authoritative workbook")
    if source.suffix.casefold() != ".xlsm" or preview.suffix.casefold() != ".xlsm":
        raise AuthoritativeCommitError("Authoritative source and preview must both be .xlsm")
    if not expected_source_sha256.strip():
        raise AuthoritativeCommitError("expected_source_sha256 is required")
    if not _has_vba(preview):
        raise AuthoritativeCommitError("Verified preview does not contain a VBA project")

    source_hash_before = file_sha256(source)
    if source_hash_before != expected_source_sha256.strip().casefold():
        raise AuthoritativeCommitError(
            "Authoritative workbook changed after preview creation; commit was refused"
        )

    preview_hash = file_sha256(preview)
    backup = Path(backup_path).resolve() if backup_path else _default_backup_path(source)
    if backup == source or backup == preview:
        raise AuthoritativeCommitError("Backup path must differ from source and preview")
    if backup.exists():
        raise AuthoritativeCommitError(f"Backup already exists: {backup}")

    backup.parent.mkdir(parents=True, exist_ok=True)
    source.parent.mkdir(parents=True, exist_ok=True)

    staging_handle = tempfile.NamedTemporaryFile(
        prefix=f".{source.stem}.commit_",
        suffix=source.suffix,
        dir=source.parent,
        delete=False,
    )
    staging = Path(staging_handle.name)
    staging_handle.close()

    source_replaced = False
    preview_removed = False
    try:
        shutil.copy2(preview, staging)
        if file_sha256(staging) != preview_hash:
            raise AuthoritativeCommitError("Staged workbook does not match verified preview")

        shutil.copy2(source, backup)
        if file_sha256(backup) != source_hash_before:
            raise AuthoritativeCommitError("Backup verification failed")

        try:
            os.replace(staging, source)
            source_replaced = True
        except PermissionError as exc:
            raise AuthoritativeCommitError(
                "Authoritative workbook is open or locked by Excel. Close it and retry the commit."
            ) from exc

        source_hash_after = file_sha256(source)
        if source_hash_after != preview_hash:
            raise AuthoritativeCommitError(
                "Authoritative workbook verification failed after replacement"
            )

        if remove_preview_after_success:
            try:
                preview.unlink()
                preview_removed = True
            except OSError:
                preview_removed = False

        return AuthoritativeCommitReport(
            source_path=str(source),
            preview_path=str(preview),
            backup_path=str(backup),
            source_sha256_before=source_hash_before,
            source_sha256_after=source_hash_after,
            preview_sha256=preview_hash,
            committed=True,
            preview_removed=preview_removed,
        )
    except Exception:
        if source_replaced and backup.exists():
            try:
                shutil.copy2(backup, source)
            except Exception:
                pass
        raise
    finally:
        if staging.exists():
            try:
                staging.unlink()
            except OSError:
                pass
