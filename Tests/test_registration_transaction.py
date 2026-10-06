from pathlib import Path

from rehab_excel.authoritative_commit import AuthoritativeCommitReport
from rehab_excel.authoritative_commit_worker import AuthoritativeCommitWorkerReport
from rehab_excel.registration_transaction import run_registration_transaction


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


def test_transaction_builds_preview_then_commits_and_reopens_without_external_close(
    tmp_path, monkeypatch
):
    source = tmp_path / "app.xlsm"
    source.write_bytes(b"source")
    events = []
    observed_preview_dirs = []

    def fake_bridge(payload):
        preview_dir = Path(str(payload["preview_dir"]))
        observed_preview_dirs.append(preview_dir)
        preview = preview_dir / "NEW_PATIENT_PREVIEW.xlsm"
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
        return _commit_report(source, Path(preview_path))

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
        }
    )

    assert [name for name, _ in events] == ["preview", "commit"]
    assert len(observed_preview_dirs) == 1
    assert observed_preview_dirs[0].parent == tmp_path / ".eka_registration_transactions"
    assert observed_preview_dirs[0] != tmp_path
    assert report.subject_key == "98"
    assert report.committed is True
    assert report.reopened is True
    assert report.close_attempts == 0


def test_transactions_use_distinct_preview_directories(tmp_path, monkeypatch):
    source = tmp_path / "app.xlsm"
    source.write_bytes(b"source")
    observed = []

    def fake_bridge(payload):
        preview_dir = Path(str(payload["preview_dir"]))
        observed.append(preview_dir)
        return {
            "subject_key": "98",
            "display_name": "TEST PATIENT",
            "output_path": str(preview_dir / "NEW_PATIENT_PREVIEW.xlsm"),
            "source_sha256_before": "abc",
        }

    monkeypatch.setattr(
        "rehab_excel.registration_transaction.run_registration_bridge", fake_bridge
    )
    monkeypatch.setattr(
        "rehab_excel.registration_transaction.commit_when_unlocked",
        lambda source_path, preview_path, **kwargs: _commit_report(
            Path(source_path), Path(preview_path)
        ),
    )

    payload = {
        "source_path": str(source),
        "preview_dir": str(tmp_path),
        "action": "new_patient",
        "values": {},
    }
    run_registration_transaction(payload)
    run_registration_transaction(payload)

    assert len(observed) == 2
    assert observed[0] != observed[1]
