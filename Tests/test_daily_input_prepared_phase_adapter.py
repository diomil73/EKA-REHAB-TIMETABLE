import json

import pytest

from rehab_excel.authoritative_commit import file_sha256
from rehab_excel.daily_input_phase_evidence import EvidenceKind
from rehab_excel.daily_input_phase_verifier_contract import (
    PhaseVerifierNotImplementedError,
    decode_phase_verifier_report,
    validate_phase_verifier_report,
)
from rehab_excel.daily_input_prepared_phase_adapter import (
    PreparedPhaseReportAdapterError,
    create_prepared_snapshot_phase_report,
)


def prepared_fixture(tmp_path):
    root = tmp_path / "evidence"
    root.mkdir()
    workbook = root / "prepared-fixture.xlsm"
    workbook.write_bytes(b"synthetic saved xlsm bytes")
    payload = tmp_path / "prepared-payload.tsv"
    payload.write_text("SYNTHETIC\\tPAYLOAD\\n", encoding="utf-8")
    handoff = {
        "status": "PREPARED_NOT_PERSISTED",
        "operation_id": "prepared-adapter-op",
        "service_date": "2026-11-25",
        "source_sha256": "a" * 64,
        "payload_sha256": file_sha256(payload),
        "existing_journal_rows": 0,
        "journal_prefix_sha256": "b" * 64,
        "prepared_row": ["prepared-adapter-op", 0, "PREPARED", "c" * 64, "payload"],
    }
    report_path = root / "prepared-report.json"
    return root, workbook, payload, handoff, report_path


def test_prepared_adapter_runs_native_readback_and_emits_exact_snapshot_bound_report(tmp_path, monkeypatch):
    import tools.verify_daily_input_journal_prepared as readback

    root, workbook, payload, handoff, report_path = prepared_fixture(tmp_path)
    calls = {}

    def fake_verify(path, received_handoff, received_payload, *, native_excel, excel_factory):
        calls.update({
            "path": path, "handoff": received_handoff, "payload": received_payload,
            "native_excel": native_excel, "excel_factory": excel_factory,
        })
        return "16.0"

    monkeypatch.setattr(readback, "verify_prepared_reopen", fake_verify)
    marker = object()
    report = create_prepared_snapshot_phase_report(
        workbook, handoff, payload, evidence_root=root, report_path=report_path,
        report_id="prepared-adapter-report", excel_factory=marker,
    )
    assert calls["path"] == workbook.resolve()
    assert calls["handoff"] is handoff
    assert calls["payload"] == payload.resolve()
    assert calls["native_excel"] is True
    assert calls["excel_factory"] is marker
    assert report.evidence_kind == EvidenceKind.PREPARED_SNAPSHOT.value
    assert report.evidence_artifact == workbook.name
    assert report.evidence_sha256 == file_sha256(workbook)
    assert report.prepared_snapshot_sha256 == file_sha256(workbook)
    assert report.source_workbook_sha256 == handoff["source_sha256"]
    assert report.native_excel_version == "16.0"
    assert report_path.is_file()
    assert decode_phase_verifier_report(report_path.read_text(encoding="utf-8").strip()) == report


def test_prepared_report_is_not_authorizing_without_trusted_attestation(tmp_path, monkeypatch):
    import tools.verify_daily_input_journal_prepared as readback

    root, workbook, payload, handoff, report_path = prepared_fixture(tmp_path)
    monkeypatch.setattr(readback, "verify_prepared_reopen", lambda *a, **k: "16.0")
    report = create_prepared_snapshot_phase_report(
        workbook, handoff, payload, evidence_root=root, report_path=report_path,
        report_id="prepared-adapter-report",
    )
    from rehab_excel.daily_input_operation import DailyInputOperation
    operation = DailyInputOperation(
        handoff["operation_id"], handoff["source_sha256"], handoff["service_date"],
    )
    with pytest.raises(PhaseVerifierNotImplementedError, match="trusted verifier attestation is not integrated"):
        validate_phase_verifier_report(report, operation)


def test_prepared_adapter_rejects_workbook_outside_evidence_root_before_verification(tmp_path, monkeypatch):
    import tools.verify_daily_input_journal_prepared as readback

    root, workbook, payload, handoff, report_path = prepared_fixture(tmp_path)
    outside = tmp_path / "outside.xlsm"
    outside.write_bytes(workbook.read_bytes())
    called = []
    monkeypatch.setattr(readback, "verify_prepared_reopen", lambda *a, **k: called.append(True))
    with pytest.raises(PreparedPhaseReportAdapterError, match="inside evidence root"):
        create_prepared_snapshot_phase_report(
            outside, handoff, payload, evidence_root=root, report_path=report_path,
            report_id="prepared-adapter-report",
        )
    assert called == []
    assert not report_path.exists()


def test_prepared_adapter_never_overwrites_an_existing_report_or_workbook(tmp_path, monkeypatch):
    import tools.verify_daily_input_journal_prepared as readback

    root, workbook, payload, handoff, report_path = prepared_fixture(tmp_path)
    report_path.write_text("KEEP EXISTING", encoding="utf-8")
    before = report_path.read_bytes()
    monkeypatch.setattr(readback, "verify_prepared_reopen", lambda *a, **k: "16.0")
    with pytest.raises(PreparedPhaseReportAdapterError, match="already exists"):
        create_prepared_snapshot_phase_report(
            workbook, handoff, payload, evidence_root=root, report_path=report_path,
            report_id="prepared-adapter-report",
        )
    assert report_path.read_bytes() == before
