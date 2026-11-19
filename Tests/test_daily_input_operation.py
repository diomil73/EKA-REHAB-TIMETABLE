from dataclasses import replace

import pytest

from rehab_excel.daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationError,
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)


def operation():
    return DailyInputOperation(
        operation_id="op-20261120-demo-001",
        source_sha256="a" * 64,
        service_date="2026-11-20",
    )


def advance(value, target, **kwargs):
    return advance_daily_input_operation(value, target, **kwargs)


def test_happy_path_requires_every_verified_stage_before_success():
    value = operation()
    for phase in (
        Phase.INPUT_AND_LEDGER_COMMITTED,
        Phase.PROJECTIONS_VERIFIED,
        Phase.REPLACEMENTS_RESOLVED,
        Phase.APPLIED_ARCHIVE_VERIFIED,
        Phase.COMPLETE,
    ):
        value = advance(value, phase)
    assert value.phase == Phase.COMPLETE
    assert value.is_terminal
    assert value.may_report_success


def test_manual_review_must_be_resolved_before_replacement_stage():
    value = advance(operation(), Phase.INPUT_AND_LEDGER_COMMITTED)
    value = advance(value, Phase.PROJECTIONS_VERIFIED)
    value = advance(value, Phase.PENDING_OPERATOR_DECISION)
    assert not value.may_report_success
    with pytest.raises(DailyInputOperationError, match="Illegal"):
        advance(value, Phase.APPLIED_ARCHIVE_VERIFIED)
    value = advance(value, Phase.REPLACEMENTS_RESOLVED, operator_resolution="manual_completed")
    assert value.operator_resolution == "manual_completed"


def test_automatic_proposal_acceptance_is_explicit():
    value = advance(operation(), Phase.INPUT_AND_LEDGER_COMMITTED)
    value = advance(value, Phase.PROJECTIONS_VERIFIED)
    value = advance(value, Phase.PENDING_OPERATOR_DECISION)
    value = advance(value, Phase.REPLACEMENTS_RESOLVED, operator_resolution="automatic_accepted")
    assert value.operator_resolution == "automatic_accepted"


@pytest.mark.parametrize("resolution", ["", "declined", "cancelled", "yes"])
def test_pending_decision_rejects_ambiguous_or_unfinished_resolution(resolution):
    value = advance(operation(), Phase.INPUT_AND_LEDGER_COMMITTED)
    value = advance(value, Phase.PROJECTIONS_VERIFIED)
    value = advance(value, Phase.PENDING_OPERATOR_DECISION)
    with pytest.raises(DailyInputOperationError, match="Operator decision"):
        advance(value, Phase.REPLACEMENTS_RESOLVED, operator_resolution=resolution)


def test_failed_operation_requires_verified_rollback_before_terminal_state():
    value = advance(operation(), Phase.FAILED, failure_code="REFRESH_FAILED", failure_detail="test")
    assert not value.is_terminal
    with pytest.raises(DailyInputOperationError, match="Illegal"):
        advance(operation(), Phase.ROLLED_BACK)
    with pytest.raises(DailyInputOperationError, match="verified rollback"):
        advance(value, Phase.ROLLED_BACK)
    value = advance(value, Phase.ROLLED_BACK, rollback_verified=True)
    assert value.rollback_verified
    assert value.is_terminal
    assert not value.may_report_success
    assert value.failure_code == "REFRESH_FAILED"


def test_failure_requires_a_reason():
    with pytest.raises(DailyInputOperationError, match="failure_code"):
        advance(operation(), Phase.FAILED)


def test_exact_transition_replay_is_idempotent():
    value = operation()
    next_value = advance(value, Phase.INPUT_AND_LEDGER_COMMITTED)
    assert advance(next_value, Phase.INPUT_AND_LEDGER_COMMITTED) is next_value
    assert next_value.revision == 1


def test_conflicting_transition_replay_is_rejected():
    value = advance(operation(), Phase.INPUT_AND_LEDGER_COMMITTED)
    with pytest.raises(DailyInputOperationError, match="Conflicting replay"):
        advance(value, Phase.INPUT_AND_LEDGER_COMMITTED, failure_code="different")


def test_revision_increments_once_per_transition():
    value = operation()
    value = advance(value, Phase.INPUT_AND_LEDGER_COMMITTED)
    value = advance(value, Phase.PROJECTIONS_VERIFIED)
    assert value.revision == 2


def test_no_success_shortcut_or_transition_out_of_terminal_state():
    value = operation()
    with pytest.raises(DailyInputOperationError, match="Illegal"):
        advance(value, Phase.COMPLETE)
    for terminal in (Phase.COMPLETE, Phase.ROLLED_BACK):
        current = operation()
        if terminal == Phase.COMPLETE:
            for phase in (
                Phase.INPUT_AND_LEDGER_COMMITTED,
                Phase.PROJECTIONS_VERIFIED,
                Phase.REPLACEMENTS_RESOLVED,
                Phase.APPLIED_ARCHIVE_VERIFIED,
                Phase.COMPLETE,
            ):
                current = advance(current, phase)
        else:
            current = advance(current, Phase.FAILED, failure_code="X")
            current = advance(current, Phase.ROLLED_BACK, rollback_verified=True)
        with pytest.raises(DailyInputOperationError, match="Terminal"):
            advance(current, Phase.FAILED, failure_code="again")


def test_operation_identity_and_source_hash_are_mandatory():
    with pytest.raises(DailyInputOperationError, match="operation_id"):
        replace(operation(), operation_id=" ")
    with pytest.raises(DailyInputOperationError, match="source_sha256"):
        replace(operation(), source_sha256="not-a-hash")


def test_pending_state_cannot_claim_resolution():
    with pytest.raises(DailyInputOperationError, match="pending decision"):
        replace(operation(), phase=Phase.PENDING_OPERATOR_DECISION, operator_resolution="manual_completed")

def test_rolled_back_requires_explicit_verified_rollback_marker():
    with pytest.raises(DailyInputOperationError, match="verified rollback evidence"):
        replace(operation(), phase=Phase.ROLLED_BACK, failure_code="X")
    with pytest.raises(DailyInputOperationError, match="only valid in ROLLED_BACK"):
        replace(operation(), rollback_verified=True)


def test_rollback_marker_cannot_be_attached_to_success_transition():
    with pytest.raises(DailyInputOperationError, match="Unexpected transition metadata"):
        advance(operation(), Phase.INPUT_AND_LEDGER_COMMITTED, rollback_verified=True)
