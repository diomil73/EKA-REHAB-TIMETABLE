"""Synthetic-only proof-obligation tests. No Excel/VBA or Apply execution."""
from dataclasses import replace

import pytest

from rehab_excel.daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_phase_evidence import (
    EvidenceKind as Proof,
    PhaseEvidenceContractError,
    required_evidence_for_transition,
)


def fresh():
    return DailyInputOperation(
        operation_id="FICTIONAL-PHASE-PROOF",
        source_sha256="a" * 64,
        service_date="2026-11-24",
    )


def next_state(previous, phase, **kwargs):
    return advance_daily_input_operation(previous, phase, **kwargs)


def test_each_straight_through_phase_has_distinct_independent_reopen_obligation():
    previous = fresh()
    for phase, proof in (
        (Phase.INPUT_AND_LEDGER_COMMITTED, Proof.INPUT_AND_LEDGER_REOPEN),
        (Phase.PROJECTIONS_VERIFIED, Proof.PROJECTIONS_REOPEN),
        (Phase.REPLACEMENTS_RESOLVED, Proof.REPLACEMENTS_REOPEN),
        (Phase.APPLIED_ARCHIVE_VERIFIED, Proof.APPLIED_ARCHIVE_REOPEN),
        (Phase.COMPLETE, Proof.FINAL_WORKBOOK_REOPEN),
    ):
        candidate = next_state(previous, phase)
        assert required_evidence_for_transition(previous, candidate) == (proof,)
        previous = candidate
    assert previous.may_report_success
    # A requirements list is not evidence that any real workbook was verified.


def test_pending_decision_requires_separate_inventory_and_operator_resolution():
    previous = next_state(fresh(), Phase.INPUT_AND_LEDGER_COMMITTED)
    previous = next_state(previous, Phase.PROJECTIONS_VERIFIED)
    pending = next_state(previous, Phase.PENDING_OPERATOR_DECISION)
    assert required_evidence_for_transition(previous, pending) == (
        Proof.PENDING_DECISION_INVENTORY,
    )
    resolved = next_state(pending, Phase.REPLACEMENTS_RESOLVED,
                          operator_resolution="manual_completed")
    assert required_evidence_for_transition(pending, resolved) == (
        Proof.OPERATOR_RESOLUTION, Proof.REPLACEMENTS_REOPEN,
    )


def test_failure_and_verified_rollback_each_require_distinct_evidence():
    original = fresh()
    failed = next_state(original, Phase.FAILED, failure_code="FICTIONAL")
    rolled = next_state(failed, Phase.ROLLED_BACK, rollback_verified=True)
    assert required_evidence_for_transition(original, failed) == (
        Proof.FAILURE_DIAGNOSTIC,
    )
    assert required_evidence_for_transition(failed, rolled) == (
        Proof.FULL_WORKBOOK_ROLLBACK,
    )
    assert not rolled.may_report_success


def test_late_failure_after_operator_decision_preserves_evidence_and_requires_full_rollback():
    state = fresh()
    for phase in (Phase.INPUT_AND_LEDGER_COMMITTED, Phase.PROJECTIONS_VERIFIED,
                  Phase.PENDING_OPERATOR_DECISION):
        state = next_state(state, phase)
    resolved = next_state(state, Phase.REPLACEMENTS_RESOLVED,
                          operator_resolution="automatic_accepted")
    failed = next_state(resolved, Phase.FAILED, failure_code="FICTIONAL_AFTER_DECISION")
    rolled = next_state(failed, Phase.ROLLED_BACK, rollback_verified=True)
    assert required_evidence_for_transition(resolved, failed) == (
        Proof.FAILURE_DIAGNOSTIC,
    )
    assert required_evidence_for_transition(failed, rolled) == (
        Proof.FULL_WORKBOOK_ROLLBACK,
    )
    assert rolled.operator_resolution == "automatic_accepted"


@pytest.mark.parametrize("mutation", [
    {"operation_id": "wrong-id"},
    {"source_sha256": "b" * 64},
    {"service_date": "2026-11-25"},
    {"revision": 9},
])
def test_transition_proposal_rejects_identity_and_revision_drift(mutation):
    previous = fresh()
    candidate = next_state(previous, Phase.INPUT_AND_LEDGER_COMMITTED)
    with pytest.raises(PhaseEvidenceContractError, match="identity or revision"):
        required_evidence_for_transition(previous, replace(candidate, **mutation))


def test_transition_proposal_rejects_shortcuts_and_noop_replay():
    previous = fresh()
    with pytest.raises(PhaseEvidenceContractError, match="non-PREPARED"):
        required_evidence_for_transition(previous, previous)
    with pytest.raises(PhaseEvidenceContractError, match="Illegal"):
        required_evidence_for_transition(
            previous, replace(previous, phase=Phase.COMPLETE, revision=1),
        )


def test_transition_proposal_rejects_dishonest_metadata_even_with_legal_phase():
    previous = fresh()
    candidate = next_state(previous, Phase.INPUT_AND_LEDGER_COMMITTED)
    with pytest.raises(PhaseEvidenceContractError, match="metadata"):
        required_evidence_for_transition(
            previous, replace(candidate, failure_code="UNRELATED"),
        )


def test_policy_is_read_only_and_does_not_authorize_apply():
    from inspect import getsource
    import rehab_excel.daily_input_phase_evidence as policy

    source = " ".join(getsource(policy).lower().split())
    assert "no trusted phase verifier is registered yet" in source
    assert "SaveCopyAs" not in source
    assert "ThisWorkbook.Save" not in source
    assert "DailyInputJournalApplyEnabled = True" not in source
