from pathlib import Path

import pytest

from rehab_excel.authoritative_commit import AuthoritativeCommitReport
from rehab_excel.authoritative_commit_worker import AuthoritativeCommitWorkerReport
from rehab_excel.registration_transaction import (
    RegistrationTransactionError,
    run_registration_transaction,
)


def _commit_report(source: Path, preview: Path) -> AuthoritativeCommitWorkerReport:
    commit = AuthoritativeCommitReport(
        source_path=str(source),
        preview_path=str(preview),
        backup_path=str(source.with_name("backup.xlsm")),
        source_sha256_before="before",
        source_sha256_after="after",
        preview_sha256="after",
        committed=True,
        preview_removed=True,
    )
    return AuthoritativeCommitWorkerReport(commit=commit, attempts=1, reopened=True)


def test_transaction_closes_before_preview_then_commits_and_reopens(tmp_path, monkeypatch):
    source = tmp_path / "app.xlsm"
    preview = tmp_path / "NEW_PATIENT_PREVIEW.xlsm"
    source.write_bytes(b"source")
    preview.write_bytes(b"preview")
    events = []

    def fake_close(path):
        events.append(("close", Path(path)))

    def fake_bridge(payload):
        events.append(("preview", Path(str(payload["source_path"]))))
        return {
            "subject_key": "98",
            "display_name": "TEST PATIENT",
            "output_path": str(preview),
            "source_sha256_before": "abc",
        }

    def fake_commit(source_path, preview_path, **kwargs):
        events.append(("commit", Path(source_path)))
        assert kwargs["expected_source_sha256"] == "abc"
        assert kwargs["reopen"] is True
        return _commit_report(source, preview)

    monkeypatch.setattr(
        "rehab_excel.registration_transaction.run_registration_bridge", fake_bridge
    )
    monkeypatch.setattr(
        "rehab_excel.registration_transaction.commit_when_unlocked", fake_commit
    )

    report = run_registration_transaction(
        {
            "source_path": str(source),
            "preview_dir": str(tmp_path),
            "action": "new_patient",
            "values": {},
            "overwrite": True,
        },
        close_delay_seconds=0,
        close_service=fake_close,
    )

    assert [name for name, _ in events] == ["close", "preview", "commit"]
    assert report.subject_key == "98"
    assert report.committed is True
    assert report.reopened is True


def test_transaction_retries_excel_close_until_vba_event_has_returned(tmp_path, monkeypatch):
    source = tmp_path / "app.xlsm"
    source.write_bytes(b"source")
    attempts = {"count": 0}
    sleeps = []

    def fake_close(_path):
        attempts["count"] += 1
        if attempts["count"] < 3:
            raise RuntimeError("Workbook.Close failed")

    monkeypatch.setattr(
        "rehab_excel.registration_transaction.run_registration_bridge",
        lambda payload: (_ for _ in ()).throw(RegistrationTransactionError("stop")),
    )

    with pytest.raises(RegistrationTransactionError, match="stop"):
        run_registration_transaction(
            {
                "source_path": str(source),
                "preview_dir": str(tmp_path),
                "action": "new_patient",
                "values": {},
            },
            close_delay_seconds=0,
            close_timeout_seconds=5,
            close_poll_seconds=0.1,
            close_service=fake_close,
            sleep_service=lambda value: sleeps.append(value),
        )

    assert attempts["count"] == 3
    assert sleeps == [0.1, 0.1]


def test_transaction_validates_close_timing(tmp_path):
    source = tmp_path / "app.xlsm"
    source.write_bytes(b"source")

    with pytest.raises(RegistrationTransactionError, match="close_timeout_seconds"):
        run_registration_transaction(
            {"source_path": str(source)},
            close_timeout_seconds=0,
        )
