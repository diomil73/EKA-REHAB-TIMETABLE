from pathlib import Path

import pytest

from rehab_excel.authoritative_commit import AuthoritativeCommitReport
from rehab_excel.authoritative_commit_bridge import (
    AuthoritativeCommitBridgeError,
    run_authoritative_commit_bridge,
)


def test_commit_bridge_maps_payload_to_service():
    calls = []

    def fake_service(
        source_path,
        preview_path,
        *,
        expected_source_sha256,
        backup_path,
        remove_preview_after_success,
    ):
        calls.append(
            (
                source_path,
                preview_path,
                expected_source_sha256,
                backup_path,
                remove_preview_after_success,
            )
        )
        return AuthoritativeCommitReport(
            source_path=str(source_path),
            preview_path=str(preview_path),
            backup_path=str(backup_path),
            source_sha256_before="before",
            source_sha256_after="after",
            preview_sha256="after",
            committed=True,
            preview_removed=remove_preview_after_success,
        )

    response = run_authoritative_commit_bridge(
        {
            "source_path": "C:/app.xlsm",
            "preview_path": "C:/preview.xlsm",
            "expected_source_sha256": "ABC123",
            "backup_path": "C:/backup.xlsm",
            "remove_preview_after_success": False,
        },
        commit_service=fake_service,
    )

    assert calls == [
        (
            "C:/app.xlsm",
            "C:/preview.xlsm",
            "ABC123",
            Path("C:/backup.xlsm"),
            False,
        )
    ]
    assert response["ok"] is True
    assert response["committed"] is True
    assert response["preview_removed"] is False


@pytest.mark.parametrize(
    ("payload", "message"),
    [
        ({}, "source_path is required"),
        ({"source_path": "a"}, "preview_path is required"),
        (
            {"source_path": "a", "preview_path": "b"},
            "expected_source_sha256 is required",
        ),
        (
            {
                "source_path": "a",
                "preview_path": "b",
                "expected_source_sha256": "c",
                "remove_preview_after_success": "yes",
            },
            "remove_preview_after_success must be a boolean",
        ),
    ],
)
def test_commit_bridge_validates_payload(payload, message):
    with pytest.raises(AuthoritativeCommitBridgeError, match=message):
        run_authoritative_commit_bridge(payload, commit_service=lambda *a, **k: None)
