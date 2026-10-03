from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from rehab_excel.authoritative_commit import (
    AuthoritativeCommitError,
    commit_verified_preview,
    file_sha256,
)


def _make_xlsm(path: Path, marker: str) -> None:
    with ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/vbaProject.bin", b"fake-vba")
        archive.writestr("marker.txt", marker)


def test_commit_promotes_preview_and_keeps_verified_backup(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    backup = tmp_path / "backup.xlsm"
    _make_xlsm(source, "source")
    _make_xlsm(preview, "preview")

    source_hash = file_sha256(source)
    preview_hash = file_sha256(preview)

    report = commit_verified_preview(
        source,
        preview,
        expected_source_sha256=source_hash,
        backup_path=backup,
    )

    assert report.committed is True
    assert report.source_sha256_before == source_hash
    assert report.source_sha256_after == preview_hash
    assert file_sha256(source) == preview_hash
    assert file_sha256(backup) == source_hash
    assert not preview.exists()
    assert report.preview_removed is True


def test_commit_refuses_if_authoritative_source_changed(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    _make_xlsm(source, "source")
    expected = file_sha256(source)
    _make_xlsm(preview, "preview")
    _make_xlsm(source, "changed")

    with pytest.raises(AuthoritativeCommitError, match="changed after preview creation"):
        commit_verified_preview(
            source,
            preview,
            expected_source_sha256=expected,
        )

    assert preview.exists()


def test_commit_requires_vba_preview(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    _make_xlsm(source, "source")
    with ZipFile(preview, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")

    with pytest.raises(AuthoritativeCommitError, match="VBA project"):
        commit_verified_preview(
            source,
            preview,
            expected_source_sha256=file_sha256(source),
        )


def test_commit_can_keep_preview_after_success(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    backup = tmp_path / "backup.xlsm"
    _make_xlsm(source, "source")
    _make_xlsm(preview, "preview")

    report = commit_verified_preview(
        source,
        preview,
        expected_source_sha256=file_sha256(source),
        backup_path=backup,
        remove_preview_after_success=False,
    )

    assert report.committed is True
    assert preview.exists()
    assert report.preview_removed is False
