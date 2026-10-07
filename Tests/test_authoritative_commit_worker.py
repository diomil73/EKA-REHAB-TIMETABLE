import inspect
from pathlib import Path

import pytest

from rehab_excel.authoritative_commit import (
    AuthoritativeCommitError,
    AuthoritativeCommitReport,
)
from rehab_excel.authoritative_commit_worker import (
    AuthoritativeCommitWorkerError,
    _default_close_workbook,
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


def test_default_close_uses_one_direct_unsaved_close_without_sentinel_macro():
    source = inspect.getsource(_default_close_workbook)
    assert "Close(SaveChanges=False)" in source
    assert "Workbooks.Add" not in source
    assert "ExitApplication" not in source
    assert "excel.Run" not in source


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

    with pytest.raises(AuthoritativeCommitWorkerError, match="close_delay_seconds"):
        commit_when_unlocked(
            source,
            preview,
            expected_source_sha256="abc",
            close_delay_seconds=-0.1,
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


def test_worker_can_close_authoritative_workbook_before_commit(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")
    preview.write_bytes(b"y")
    events = []

    def fake_close(path):
        events.append(("close", path))

    def fake_commit(*args, **kwargs):
        events.append(("commit", Path(args[0])))
        return _report(source, preview)

    result = commit_when_unlocked(
        source,
        preview,
        expected_source_sha256="abc",
        close_open_workbook=True,
        close_delay_seconds=1.5,
        close_service=fake_close,
        commit_service=fake_commit,
        sleep_service=lambda value: events.append(("sleep", value)),
    )

    resolved = source.resolve()
    assert events == [("sleep", 1.5), ("close", resolved), ("commit", resolved)]
    assert result.commit.committed is True


def test_worker_wraps_controlled_close_failure(tmp_path):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")
    preview.write_bytes(b"y")

    def fail_close(_path):
        raise RuntimeError("Excel refused close")

    with pytest.raises(AuthoritativeCommitWorkerError, match="Could not close authoritative workbook"):
        commit_when_unlocked(
            source,
            preview,
            expected_source_sha256="abc",
            close_open_workbook=True,
            close_delay_seconds=0,
            close_service=fail_close,
            commit_service=lambda *a, **k: _report(source, preview),
        )
