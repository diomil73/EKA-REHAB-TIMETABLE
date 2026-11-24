"""Synthetic-only non-PREPARED proposal, stale-source and reopen regressions."""
from datetime import date
import json

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.authoritative_commit import file_sha256
from rehab_excel.daily_input_journal import (
    JOURNAL_HEADERS, JOURNAL_SHEET, _row,
    append_operation_journal, decode_operation_journal,
)
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_phase_evidence import EvidenceKind
from tools import prepare_daily_input_journal_transition as transition


def fixture(tmp_path, *, history=None):
    source = tmp_path / "FICTIONAL_PREPARED.xlsm"
    first = DailyInputOperation(
        operation_id="FICTIONAL-OPERATION", source_sha256="a" * 64,
        service_date="2026-11-24",
    )
    if history is None:
        history = append_operation_journal((), first)
    workbook = Workbook()
    daily = workbook.active
    daily.title = "DAILY_INPUT"
    daily["B2"] = date(2026, 11, 24)
    journal = workbook.create_sheet(JOURNAL_SHEET)
    journal.append(JOURNAL_HEADERS)
    for row in history:
        journal.append(row)
    journal.sheet_state = "veryHidden"
    journal.protection.sheet = True
    workbook.save(source)
    workbook.close()
    return source, first, history


def append_on_synthetic_copy(source, next_state):
    """Test fixture only. Operational Excel must use verified VBA writer."""
    wb = load_workbook(source, keep_vba=True)
    wb[JOURNAL_SHEET].append(_row(next_state))
    wb.save(source)
    wb.close()


def test_proposal_contains_canonical_next_revision_but_no_write_authorization(tmp_path):
    source, first, rows = fixture(tmp_path)
    before = source.read_bytes()
    proposal = transition.prepare_transition_proposal(
        source, Phase.INPUT_AND_LEDGER_COMMITTED,
    )
    assert proposal["status"] == "TRANSITION_PROPOSAL_ONLY"
    assert proposal["operation_id"] == first.operation_id
    assert proposal["operation_source_sha256"] == first.source_sha256
    assert proposal["workbook_sha256_before"] == file_sha256(source)
    assert proposal["existing_journal_rows"] == 1
    assert proposal["required_evidence"] == [EvidenceKind.INPUT_AND_LEDGER_REOPEN.value]
    assert tuple(proposal["proposed_row"]) == _row(
        advance_daily_input_operation(first, Phase.INPUT_AND_LEDGER_COMMITTED),
    )
    assert source.read_bytes() == before
    assert set(tmp_path.iterdir()) == {source}
    transition.verify_transition_source(source, proposal)


def test_independent_reopen_requires_actual_append_and_exact_history(tmp_path):
    source, first, rows = fixture(tmp_path)
    proposal = transition.prepare_transition_proposal(
        source, Phase.INPUT_AND_LEDGER_COMMITTED,
    )
    with pytest.raises(transition.TransitionProposalError, match="count"):
        transition.verify_transition_reopen(source, proposal)
    candidate = advance_daily_input_operation(first, Phase.INPUT_AND_LEDGER_COMMITTED)
    append_on_synthetic_copy(source, candidate)
    after = source.read_bytes()
    transition.verify_transition_reopen(source, proposal)
    assert source.read_bytes() == after
    with pytest.raises(transition.TransitionProposalError, match="Workbook changed"):
        transition.verify_transition_source(source, proposal)


def test_changing_source_after_proposal_fails_closed(tmp_path):
    source, _, _ = fixture(tmp_path)
    proposal = transition.prepare_transition_proposal(source, Phase.FAILED, failure_code="FICTIONAL")
    wb = load_workbook(source, keep_vba=True)
    wb["DAILY_INPUT"]["C3"] = "UNEXPECTED SYNTHETIC CHANGE"
    wb.save(source)
    wb.close()
    with pytest.raises(transition.TransitionProposalError, match="Workbook changed"):
        transition.verify_transition_source(source, proposal)


def test_forced_phase_and_forged_proof_requirements_are_rejected(tmp_path):
    source, first, _ = fixture(tmp_path)
    proposal = transition.prepare_transition_proposal(source, Phase.INPUT_AND_LEDGER_COMMITTED)
    forged = dict(proposal, required_evidence=["caller_says_success"])
    with pytest.raises(transition.TransitionProposalError, match="proof obligations differ"):
        transition.verify_transition_source(source, forged)
    forged = dict(proposal, target_phase="COMPLETE")
    with pytest.raises(transition.TransitionProposalError, match="Proposed transition"):
        transition.verify_transition_source(source, forged)
    with pytest.raises(transition.TransitionProposalError, match="Illegal journal"):
        transition.prepare_transition_proposal(source, Phase.COMPLETE)


def test_pending_operator_choice_needs_two_distinct_obligations(tmp_path):
    first = DailyInputOperation("FICTIONAL-OPERATION", "a" * 64, "2026-11-24")
    states = [first]
    for phase in (Phase.INPUT_AND_LEDGER_COMMITTED, Phase.PROJECTIONS_VERIFIED,
                  Phase.PENDING_OPERATOR_DECISION):
        states.append(advance_daily_input_operation(states[-1], phase))
    rows = ()
    for state in states:
        rows = append_operation_journal(rows, state)
    source, _, _ = fixture(tmp_path, history=rows)
    with pytest.raises(transition.TransitionProposalError, match="Illegal"):
        transition.prepare_transition_proposal(source, Phase.REPLACEMENTS_RESOLVED)
    proposal = transition.prepare_transition_proposal(
        source, Phase.REPLACEMENTS_RESOLVED, operator_resolution="manual_completed",
    )
    assert proposal["required_evidence"] == [
        EvidenceKind.OPERATOR_RESOLUTION.value,
        EvidenceKind.REPLACEMENTS_REOPEN.value,
    ]
    transition.verify_transition_source(source, proposal)


@pytest.mark.parametrize("phase,kwargs", [
    (Phase.PREPARED, {}),
    (Phase.ROLLED_BACK, {"rollback_verified": True}),
    (Phase.FAILED, {}),
    (Phase.APPLIED_ARCHIVE_VERIFIED, {}),
])
def test_unpermitted_phase_or_missing_failure_evidence_rejected(tmp_path, phase, kwargs):
    source, _, _ = fixture(tmp_path)
    with pytest.raises(transition.TransitionProposalError):
        transition.prepare_transition_proposal(source, phase, **kwargs)


def test_reopen_refuses_an_extra_revision_and_changed_original_history(tmp_path):
    source, first, _ = fixture(tmp_path)
    proposal = transition.prepare_transition_proposal(source, Phase.FAILED, failure_code="FICTIONAL")
    failed = advance_daily_input_operation(first, Phase.FAILED, failure_code="FICTIONAL")
    append_on_synthetic_copy(source, failed)
    wb = load_workbook(source, keep_vba=True)
    wb[JOURNAL_SHEET]["A2"] = "TAMPERED FICTIONAL ID"
    wb.save(source)
    wb.close()
    with pytest.raises(transition.TransitionProposalError):
        transition.verify_transition_reopen(source, proposal)


def test_native_reopen_proof_is_used_only_when_explicitly_requested(tmp_path, monkeypatch):
    source, first, _ = fixture(tmp_path)
    proposal = transition.prepare_transition_proposal(source, Phase.INPUT_AND_LEDGER_COMMITTED)
    append_on_synthetic_copy(
        source, advance_daily_input_operation(first, Phase.INPUT_AND_LEDGER_COMMITTED),
    )
    calls = []
    monkeypatch.setattr(
        transition, "verify_native_excel_readonly",
        lambda path, rows, **kwargs: calls.append((path, rows)),
    )
    transition.verify_transition_reopen(source, proposal, native_excel=True)
    assert len(calls) == 1
    assert calls[0][0] == source
    assert len(calls[0][1]) == 2


def test_cli_refuses_overwriting_proposal_and_does_not_write_workbook(tmp_path, capsys):
    source, _, _ = fixture(tmp_path)
    out = tmp_path / "FICTIONAL_TRANSITION_PROPOSAL.json"
    before = source.read_bytes()
    assert transition.main([
        "--input", str(source), "--prepare", "INPUT_AND_LEDGER_COMMITTED",
        "--output", str(out),
    ]) == 0
    assert "NOT AUTHORIZED" in capsys.readouterr().out
    assert source.read_bytes() == before
    proposal = json.loads(out.read_text(encoding="utf-8"))
    assert proposal["status"] == transition.PROPOSAL_STATUS
    assert transition.main([
        "--input", str(source), "--prepare", "INPUT_AND_LEDGER_COMMITTED",
        "--output", str(out),
    ]) == 2
    assert "SAFETY STOP" in capsys.readouterr().out
    assert source.read_bytes() == before


def test_cli_verifies_source_and_refuses_unsaved_reopen(tmp_path, capsys):
    source, _, _ = fixture(tmp_path)
    proposal = transition.prepare_transition_proposal(source, Phase.FAILED, failure_code="FICTIONAL")
    path = tmp_path / "FICTIONAL_PROPOSAL.json"
    path.write_text(json.dumps(proposal), encoding="utf-8")
    assert transition.main(["--input", str(source), "--verify-source", str(path)]) == 0
    assert "NOT AUTHORIZED" in capsys.readouterr().out
    assert transition.main(["--input", str(source), "--verify-reopen", str(path)]) == 2
    assert "SAFETY STOP" in capsys.readouterr().out


def test_cli_failure_prints_no_sensitive_details(tmp_path, capsys):
    source, _, _ = fixture(tmp_path)
    output = tmp_path / "private-name.json"
    assert transition.main([
        "--input", str(source), "--prepare", "FAILED",
        "--failure-code", "PRIVATE_PATIENT_NAME", "--output", str(output),
    ]) == 0
    capsys.readouterr()
    assert transition.main([
        "--input", str(source), "--verify-reopen", str(output),
    ]) == 2
    log = capsys.readouterr().out
    assert "SAFETY STOP" in log
    assert "PRIVATE_PATIENT_NAME" not in log
    assert "private-name" not in log
