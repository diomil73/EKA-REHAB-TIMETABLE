from pathlib import Path

import pytest

from rehab_excel.authoritative_commit import (
    AuthoritativeCommitError,
    AuthoritativeCommitReport,
)
from rehab_excel.authoritative_commit_worker import (
    AuthoritativeCommitWorkerError,
    commit_when_unlocked,
)


def _report(source: Path, preview: Path) -> AuthoritativeCommitReport:
    return AuthoritativeCommitReport(
        source_path=str(source),
        preview_path=str(preview),
        backup_path=str(source.with_name("backup.xlsm")),
        source_sha256_before="before",
        source_sha256_after="after",
        preview_sha256="after",
        committed=True,
        preview_removed=True,
    )


def test_worker_retries_excel_lock_then_commits(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")
    preview.write_bytes(b"y")

    attempts = {"count": 0}
    sleeps = []

    def fake_commit(*args, **kwargs):
        attempts["count"] += 1
        if attempts["count"] < 3:
            raise AuthoritativeCommitError(
                "Authoritative workbook is open or locked by Excel. Close it and retry the commit."
            )
        return _report(source, preview)

    result = commit_when_unlocked(
        source,
        preview,
        expected_source_sha256="abc",
        timeout_seconds=5,
        poll_seconds=0.1,
        commit_service=fake_commit,
        sleep_service=lambda value: sleeps.append(value),
    )

    assert result.commit.committed is True
    assert result.attempts == 3
    assert sleeps == [0.1, 0.1]
    assert result.reopened is False


def test_worker_does_not_retry_non_lock_failure(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")
    preview.write_bytes(b"y")

    def fake_commit(*args, **kwargs):
        raise AuthoritativeCommitError(
            "Authoritative workbook changed after preview creation; commit was refused"
        )

    with pytest.raises(AuthoritativeCommitWorkerError, match="changed after preview"):
        commit_when_unlocked(
            source,
            preview,
            expected_source_sha256="abc",
            commit_service=fake_commit,
        )


def test_worker_reopens_after_success(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")
    preview.write_bytes(b"y")
    opened = []

    result = commit_when_unlocked(
        source,
        preview,
        expected_source_sha256="abc",
        reopen=True,
        commit_service=lambda *a, **k: _report(source, preview),
        open_service=lambda path: opened.append(path),
    )

    assert opened == [source.resolve()]
    assert result.reopened is True


def test_worker_validates_timing_arguments(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")
    preview.write_bytes(b"y")

    with pytest.raises(AuthoritativeCommitWorkerError, match="timeout_seconds"):
        commit_when_unlocked(
            source,
            preview,
            expected_source_sha256="abc",
            timeout_seconds=0,
        )

    with pytest.raises(AuthoritativeCommitWorkerError, match="poll_seconds"):
        commit_when_unlocked(
            source,
            preview,
            expected_source_sha256="abc",
            poll_seconds=0,
        )



def test_worker_writes_post_reopen_target_before_open(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")
    preview.write_bytes(b"y")
    observed = []

    def fake_open(path):
        marker = path.with_name(path.name + ".eka_next_sheet")
        observed.append(marker.read_text(encoding="utf-8"))

    result = commit_when_unlocked(
        source,
        preview,
        expected_source_sha256="abc",
        reopen=True,
        post_reopen_target="REPLACEMENTS",
        commit_service=lambda *a, **k: _report(source, preview),
        open_service=fake_open,
    )

    assert observed == ["REPLACEMENTS"]
    assert result.reopened is True
