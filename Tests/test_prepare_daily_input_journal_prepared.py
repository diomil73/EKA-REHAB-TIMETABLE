"""Fictional, non-VBA-only tests of the read-only PREPARED handoff gate."""
import json
from datetime import date

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.authoritative_commit import file_sha256
from rehab_excel.daily_input_journal import (
    JOURNAL_HEADERS, JOURNAL_SHEET, append_operation_journal, decode_operation_journal,
)
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from tools import prepare_daily_input_journal_prepared as bridge
from tools import verify_daily_input_journal_prepared as readback


def fixture(tmp_path, *, history=(), protected=True, visibility="veryHidden"):
    source = tmp_path / "fictional.xlsm"
    payload = tmp_path / "fictional-plan.tsv"
    response = tmp_path / "fictional-response.json"
    workbook = Workbook()
    daily = workbook.active
    daily.title = "DAILY_INPUT"
    daily["B2"] = date(2026, 11, 23)
    journal = workbook.create_sheet(JOURNAL_SHEET)
    for index, header in enumerate(JOURNAL_HEADERS, start=1):
        journal.cell(1, index, header)
    for row in history:
        journal.append(row)
    journal.sheet_state = visibility
    journal.protection.sheet = protected
    workbook.save(source)
    workbook.close()
    payload.write_text("PATCH\tFICTIONAL\n", encoding="utf-8")
    response.write_text(json.dumps({
        "ok": True,
        "source_unchanged": True,
        "date": "2026-11-23",
        "source_sha256_before": file_sha256(source),
        "payload_path": str(payload.resolve()),
    }), encoding="utf-8")
    return source, response, payload


def _terminal_history():
    first = DailyInputOperation("old-fixture", "a" * 64, "2026-11-20")
    rows = append_operation_journal((), first)
    for phase in (
        Phase.INPUT_AND_LEDGER_COMMITTED, Phase.PROJECTIONS_VERIFIED,
        Phase.REPLACEMENTS_RESOLVED, Phase.APPLIED_ARCHIVE_VERIFIED, Phase.COMPLETE,
    ):
        first = advance_daily_input_operation(first, phase)
        rows = append_operation_journal(rows, first)
    return rows


def test_prepared_handoff_matches_source_and_canonical_append_without_workbook_edits(tmp_path):
    source, response, payload = fixture(tmp_path)
    before = source.read_bytes()
    operation_id = "fictional-next"
    result = bridge.prepare_prepared_handoff(
        source, response, payload, operation_id=operation_id,
    )
    assert result["status"] == "PREPARED_NOT_PERSISTED"
    assert result["operation_id"] == operation_id
    assert result["service_date"] == "2026-11-23"
    assert result["source_sha256"] == file_sha256(source)
    assert result["payload_sha256"] == file_sha256(payload)
    assert result["existing_journal_rows"] == 0
    appended = decode_operation_journal([tuple(result["prepared_row"])])
    assert len(appended) == 1
    assert appended[0].phase == Phase.PREPARED
    assert appended[0].revision == 0
    assert appended[0].operation_id == operation_id
    assert source.read_bytes() == before
    assert set(tmp_path.iterdir()) == {source, response, payload}


def test_prepared_handoff_accepts_a_new_unique_operation_after_terminal_history(tmp_path):
    history = _terminal_history()
    source, response, payload = fixture(tmp_path, history=history)
    before = source.read_bytes()
    result = bridge.prepare_prepared_handoff(
        source, response, payload, operation_id="fictional-second",
    )
    full = decode_operation_journal([*history, tuple(result["prepared_row"])])
    assert result["existing_journal_rows"] == len(history)
    assert full[-1].operation_id == "fictional-second"
    assert full[-1].revision == 0
    assert full[-2].phase == Phase.COMPLETE
    assert source.read_bytes() == before


def test_unresolved_operation_blocks_new_prepared_without_creating_a_file(tmp_path):
    first = DailyInputOperation("pending-fixture", "a" * 64, "2026-11-20")
    failed = advance_daily_input_operation(first, Phase.FAILED, failure_code="FICTIONAL")
    history = append_operation_journal(append_operation_journal((), first), failed)
    source, response, payload = fixture(tmp_path, history=history)
    before = source.read_bytes()
    with pytest.raises(bridge.PreparedJournalPreflightError, match="rejected"):
        bridge.prepare_prepared_handoff(
            source, response, payload, operation_id="fictional-second",
        )
    assert source.read_bytes() == before


@pytest.mark.parametrize("visibility,protected", [
    ("visible", True), ("veryHidden", False),
])
def test_requires_existing_protected_veryhidden_journal(tmp_path, visibility, protected):
    source, response, payload = fixture(
        tmp_path, visibility=visibility, protected=protected,
    )
    with pytest.raises(bridge.PreparedJournalPreflightError, match="protection"):
        bridge.prepare_prepared_handoff(source, response, payload)


def test_rejects_existing_corrupt_journal_without_modification(tmp_path):
    history = _terminal_history()
    changed = list(history[0])
    changed[3] = "0" * 64
    source, response, payload = fixture(
        tmp_path, history=(tuple(changed), *history[1:]),
    )
    before = source.read_bytes()
    with pytest.raises(bridge.PreparedJournalPreflightError, match="history invalid"):
        bridge.prepare_prepared_handoff(source, response, payload)
    assert source.read_bytes() == before


@pytest.mark.parametrize("changed_key,replacement,expected", [
    ("source_sha256_before", "0" * 64, "Workbook changed"),
    ("date", "2026-11-24", "service date mismatch"),
    ("ok", False, "receipt not verified"),
    ("source_unchanged", False, "receipt not verified"),
    ("payload_path", "different.tsv", "payload path mismatch"),
])
def test_fails_closed_on_stale_or_unverified_preparation(
    tmp_path, changed_key, replacement, expected,
):
    source, response, payload = fixture(tmp_path)
    original = source.read_bytes()
    receipt = json.loads(response.read_text(encoding="utf-8"))
    receipt[changed_key] = replacement
    response.write_text(json.dumps(receipt), encoding="utf-8")
    with pytest.raises(bridge.PreparedJournalPreflightError, match=expected):
        bridge.prepare_prepared_handoff(source, response, payload)
    assert source.read_bytes() == original


def test_cli_creates_only_new_handoff_metadata_not_an_excel_or_journal_write(tmp_path, capsys):
    source, response, payload = fixture(tmp_path)
    before = source.read_bytes()
    out = tmp_path / "prepared-fictitious.json"
    args = [
        "--input", str(source), "--preparation-response", str(response),
        "--payload", str(payload), "--output", str(out),
    ]
    assert bridge.main(args) == 0
    assert "NOT PERSISTED" in capsys.readouterr().out
    assert out.is_file()
    metadata = json.loads(out.read_text(encoding="utf-8"))
    assert metadata["status"] == "PREPARED_NOT_PERSISTED"
    assert metadata["existing_journal_rows"] == 0
    assert source.read_bytes() == before
    old = out.read_bytes()
    assert bridge.main(args) == 2
    assert "SAFETY STOP" in capsys.readouterr().out
    assert out.read_bytes() == old


def test_failed_cli_has_no_output_or_leaked_sensitive_detail(tmp_path, capsys):
    source, response, payload = fixture(tmp_path)
    content = json.loads(response.read_text(encoding="utf-8"))
    content["payload_path"] = "PRIVATE PATIENT NAME.tsv"
    response.write_text(json.dumps(content), encoding="utf-8")
    out = tmp_path / "should-not-exist.json"
    args = [
        "--input", str(source), "--preparation-response", str(response),
        "--payload", str(payload), "--output", str(out),
    ]
    assert bridge.main(args) == 2
    printed = capsys.readouterr().out
    assert "PRIVATE" not in printed
    assert "SAFETY STOP" in printed
    assert not out.exists()



def _append_proposal_to_synthetic_book(source, handoff):
    """Test fixture only. Real caller must use guarded native VBA+Save."""
    workbook = load_workbook(source, keep_vba=True)
    ws = workbook[JOURNAL_SHEET]
    ws.append(handoff["prepared_row"])
    workbook.save(source)
    workbook.close()


def test_independent_prepared_readback_requires_durable_synthetic_append(tmp_path, capsys):
    source, response, payload = fixture(tmp_path)
    prepared = bridge.prepare_prepared_handoff(
        source, response, payload, operation_id="fixture-proposal",
    )
    with pytest.raises(readback.PreparedJournalReadbackError, match="row count"):
        readback.verify_prepared_reopen(source, prepared)
    _append_proposal_to_synthetic_book(source, prepared)
    source_after = source.read_bytes()
    readback.verify_prepared_reopen(source, prepared)
    assert source.read_bytes() == source_after
    handoff_path = tmp_path / "handoff.json"
    handoff_path.write_text(json.dumps(prepared), encoding="utf-8")
    assert readback.main(["--input", str(source), "--handoff", str(handoff_path)]) == 0
    assert "NO APPLY/RESTORE APPROVAL" in capsys.readouterr().out
    assert source.read_bytes() == source_after


def test_independent_readback_detects_corrupted_previous_history_even_if_valid(tmp_path):
    history = _terminal_history()
    source, response, payload = fixture(tmp_path, history=history)
    proposal = bridge.prepare_prepared_handoff(
        source, response, payload, operation_id="next-fixture",
    )
    # Rebuild a different but internally valid prior history.
    old = DailyInputOperation("old-fixture", "b" * 64, "2026-11-20")
    mutated = append_operation_journal((), old)
    for phase in (
        Phase.INPUT_AND_LEDGER_COMMITTED, Phase.PROJECTIONS_VERIFIED,
        Phase.REPLACEMENTS_RESOLVED, Phase.APPLIED_ARCHIVE_VERIFIED, Phase.COMPLETE,
    ):
        old = advance_daily_input_operation(old, phase)
        mutated = append_operation_journal(mutated, old)
    workbook = load_workbook(source, keep_vba=True)
    ws = workbook[JOURNAL_SHEET]
    for idx, row in enumerate(mutated, 2):
        for col, value in enumerate(row, 1):
            ws.cell(idx, col).value = value
    ws.append(proposal["prepared_row"])
    workbook.save(source)
    workbook.close()
    with pytest.raises(readback.PreparedJournalReadbackError, match="Historical journal prefix"):
        readback.verify_prepared_reopen(source, proposal)


@pytest.mark.parametrize("mutation", ["extra-row", "wrong-proposed-row", "visible", "date"])
def test_independent_readback_stops_on_unverified_reopen(tmp_path, mutation):
    source, response, payload = fixture(tmp_path)
    proposal = bridge.prepare_prepared_handoff(
        source, response, payload, operation_id="fixture-proposal",
    )
    _append_proposal_to_synthetic_book(source, proposal)
    workbook = load_workbook(source, keep_vba=True)
    journal = workbook[JOURNAL_SHEET]
    if mutation == "extra-row":
        first = DailyInputOperation("fixture-proposal", proposal["source_sha256"], "2026-11-23")
        failed = advance_daily_input_operation(first, Phase.FAILED, failure_code="FICTIONAL")
        from rehab_excel.daily_input_journal import _row
        journal.append(_row(failed))
    elif mutation == "wrong-proposed-row":
        journal.cell(2, 3).value = "COMPLETE"
    elif mutation == "visible":
        journal.sheet_state = "visible"
    else:
        workbook["DAILY_INPUT"]["B2"] = date(2026, 11, 24)
    workbook.save(source)
    workbook.close()
    with pytest.raises(readback.PreparedJournalReadbackError):
        readback.verify_prepared_reopen(source, proposal)


def test_readback_rejects_forged_metadata_and_does_not_echo_patient_data(tmp_path, capsys):
    source, response, payload = fixture(tmp_path)
    proposal = bridge.prepare_prepared_handoff(
        source, response, payload, operation_id="fixture-proposal",
    )
    _append_proposal_to_synthetic_book(source, proposal)
    altered = dict(proposal, operation_id="PRIVATE PATIENT")
    with pytest.raises(readback.PreparedJournalReadbackError, match="Prepared record differs"):
        readback.verify_prepared_reopen(source, altered)
    handoff_path = tmp_path / "altered.json"
    handoff_path.write_text(json.dumps(altered), encoding="utf-8")
    assert readback.main(["--input", str(source), "--handoff", str(handoff_path)]) == 2
    output = capsys.readouterr().out
    assert "SAFETY STOP" in output
    assert "PRIVATE" not in output
