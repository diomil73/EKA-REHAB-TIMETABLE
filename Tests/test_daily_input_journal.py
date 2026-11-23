import pytest

from rehab_excel.daily_input_operation import DailyInputOperation, DailyInputOperationPhase as Phase, advance_daily_input_operation
from rehab_excel.daily_input_journal import (
    JOURNAL_HEADERS,
    DailyInputJournalError,
    append_operation_journal,
    decode_operation_journal,
    verify_operation_journal_readback,
    _row,
)


def op():
    return DailyInputOperation("op-1", "a" * 64, "2026-11-20")


def test_journal_schema_is_fixed():
    assert JOURNAL_HEADERS == ("OperationID", "Revision", "Phase", "PayloadSHA256", "Payload")


def test_append_only_history_round_trips_every_transition():
    states = [op()]
    for phase in (Phase.INPUT_AND_LEDGER_COMMITTED, Phase.PROJECTIONS_VERIFIED,
                  Phase.REPLACEMENTS_RESOLVED, Phase.APPLIED_ARCHIVE_VERIFIED, Phase.COMPLETE):
        states.append(advance_daily_input_operation(states[-1], phase))
    rows = ()
    for state in states:
        rows = append_operation_journal(rows, state)
    assert decode_operation_journal(rows) == tuple(states)


def test_exact_append_replay_is_idempotent():
    first = op()
    rows = append_operation_journal((), first)
    assert append_operation_journal(rows, first) == rows


def test_rejects_non_prepared_first_entry():
    state = advance_daily_input_operation(op(), Phase.INPUT_AND_LEDGER_COMMITTED)
    with pytest.raises(DailyInputJournalError, match="First journal entry"):
        append_operation_journal((), state)


def test_rejects_missing_revision():
    first = op()
    committed = advance_daily_input_operation(first, Phase.INPUT_AND_LEDGER_COMMITTED)
    third = advance_daily_input_operation(committed, Phase.PROJECTIONS_VERIFIED)
    rows = append_operation_journal((), first) + (_row(third),)
    with pytest.raises(DailyInputJournalError, match="revision"):
        decode_operation_journal(rows)


def test_rejects_checksum_corruption():
    row = list(append_operation_journal((), op())[0])
    row[4] += "x"
    with pytest.raises(DailyInputJournalError, match="checksum"):
        decode_operation_journal((tuple(row),))


def test_rejects_row_payload_metadata_mismatch():
    row = list(append_operation_journal((), op())[0])
    row[2] = Phase.COMPLETE.value
    with pytest.raises(DailyInputJournalError, match="metadata mismatch"):
        decode_operation_journal((tuple(row),))


def test_rejects_mixed_operation_identity():
    first = op()
    rows = append_operation_journal((), first)
    other = DailyInputOperation("op-2", "b" * 64, "2026-11-20")
    with pytest.raises(DailyInputJournalError, match="different operation"):
        append_operation_journal(rows, other)


def test_readback_checks_complete_history_exactly():
    rows = append_operation_journal((), op())
    assert verify_operation_journal_readback(rows, rows) == (op(),)
    with pytest.raises(DailyInputJournalError, match="readback"):
        verify_operation_journal_readback((), rows)


def test_journal_requires_valid_rollback_state():
    failed = advance_daily_input_operation(op(), Phase.FAILED, failure_code="REFRESH")
    rolled = advance_daily_input_operation(failed, Phase.ROLLED_BACK, rollback_verified=True)
    rows = append_operation_journal((), op())
    rows = append_operation_journal(rows, failed)
    rows = append_operation_journal(rows, rolled)
    assert decode_operation_journal(rows)[-1].rollback_verified

def test_rejects_malformed_payload_field_types():
    import json
    from hashlib import sha256

    row = list(append_operation_journal((), op())[0])
    payload = json.loads(row[4].split(":", 1)[1])
    payload["service_date"] = ["not", "a", "string"]
    raw = json.dumps(payload, ensure_ascii=True, allow_nan=False,
                     sort_keys=True, separators=(",", ":"))
    row[3] = sha256(("diomil-journal-v1:" + raw).encode("ascii")).hexdigest()
    row[4] = "diomil-journal-v1:" + raw
    with pytest.raises(DailyInputJournalError, match="payload"):
        decode_operation_journal((tuple(row),))


def _completed(operation):
    states = [operation]
    for phase in (
        Phase.INPUT_AND_LEDGER_COMMITTED,
        Phase.PROJECTIONS_VERIFIED,
        Phase.REPLACEMENTS_RESOLVED,
        Phase.APPLIED_ARCHIVE_VERIFIED,
        Phase.COMPLETE,
    ):
        states.append(advance_daily_input_operation(states[-1], phase))
    return tuple(states)


def test_sequential_completed_operations_keep_full_original_history():
    prior = _completed(op())
    next_start = DailyInputOperation("op-2", "b" * 64, "2026-11-23")
    next_end = _completed(next_start)
    rows = ()
    for state in (*prior, *next_end):
        rows = append_operation_journal(rows, state)
    assert decode_operation_journal(rows) == (*prior, *next_end)
    assert verify_operation_journal_readback(rows, rows) == (*prior, *next_end)
    assert append_operation_journal(rows, next_end[-1]) == rows


def test_second_operation_allowed_after_verified_rolled_back():
    first = op()
    failed = advance_daily_input_operation(first, Phase.FAILED, failure_code="FICTIONAL")
    rolled = advance_daily_input_operation(failed, Phase.ROLLED_BACK, rollback_verified=True)
    rows = ()
    for state in (first, failed, rolled):
        rows = append_operation_journal(rows, state)
    second = DailyInputOperation("op-2", "b" * 64, "2026-11-24")
    rows = append_operation_journal(rows, second)
    assert decode_operation_journal(rows) == (first, failed, rolled, second)


@pytest.mark.parametrize("terminal_gap", ["PREPARED", "FAILED", "PENDING_OPERATOR_DECISION"])
def test_second_operation_blocked_while_prior_unresolved(terminal_gap):
    first = op()
    if terminal_gap == "FAILED":
        first = advance_daily_input_operation(first, Phase.FAILED, failure_code="FICTIONAL")
    elif terminal_gap == "PENDING_OPERATOR_DECISION":
        first = advance_daily_input_operation(first, Phase.INPUT_AND_LEDGER_COMMITTED)
        first = advance_daily_input_operation(first, Phase.PROJECTIONS_VERIFIED)
        first = advance_daily_input_operation(first, Phase.PENDING_OPERATOR_DECISION)
    other = DailyInputOperation("op-2", "b" * 64, "2026-11-24")
    with pytest.raises(DailyInputJournalError, match="different operation|prior terminal"):
        append_operation_journal((_row(first),) if terminal_gap == "PREPARED"
                                 else _rows_to_state(first), other)


def _rows_to_state(target):
    """Create the exact legal prefix needed for a nonterminal test."""
    rows = append_operation_journal((), op())
    if target.phase == Phase.FAILED:
        return append_operation_journal(rows, target)
    if target.phase == Phase.PENDING_OPERATOR_DECISION:
        states = _completed(op())
        for state in states[1:3]:
            rows = append_operation_journal(rows, state)
        return append_operation_journal(
            rows, advance_daily_input_operation(states[2], Phase.PENDING_OPERATOR_DECISION)
        )
    return rows


def test_operation_id_cannot_be_reused_after_other_completed_operation():
    first = _completed(op())
    second = _completed(DailyInputOperation("op-2", "b" * 64, "2026-11-24"))
    rows = ()
    for state in (*first, *second):
        rows = append_operation_journal(rows, state)
    with pytest.raises(DailyInputJournalError, match="cannot be reused"):
        append_operation_journal(rows, op())
    with pytest.raises(DailyInputJournalError, match="cannot be reused"):
        decode_operation_journal((*rows, _row(op())))


def test_new_operation_requires_prepared_revision_zero_and_no_history_edit():
    rows = ()
    for state in _completed(op()):
        rows = append_operation_journal(rows, state)
    next_op = DailyInputOperation("op-2", "b" * 64, "2026-11-24")
    later = advance_daily_input_operation(next_op, Phase.INPUT_AND_LEDGER_COMMITTED)
    with pytest.raises(DailyInputJournalError, match="PREPARED revision 0"):
        append_operation_journal(rows, later)
    with pytest.raises(DailyInputJournalError, match="PREPARED revision 0"):
        decode_operation_journal((*rows, _row(later)))
    assert decode_operation_journal(rows) == _completed(op())
    with pytest.raises(DailyInputJournalError, match="readback"):
        verify_operation_journal_readback(rows, (*rows, _row(next_op)))


def test_sequence_decoder_rejects_unterminated_transition_even_after_recovery():
    first = op()
    failed = advance_daily_input_operation(first, Phase.FAILED, failure_code="FICTIONAL")
    next_op = DailyInputOperation("op-2", "b" * 64, "2026-11-24")
    with pytest.raises(DailyInputJournalError, match="prior terminal"):
        decode_operation_journal((_row(first), _row(failed), _row(next_op)))

@pytest.mark.parametrize("resolution", ["manual_completed", "automatic_accepted"])
def test_journal_replays_late_failure_and_rollback_preserving_operator_decision(resolution):
    states = [op()]
    for phase in (
        Phase.INPUT_AND_LEDGER_COMMITTED,
        Phase.PROJECTIONS_VERIFIED,
        Phase.PENDING_OPERATOR_DECISION,
    ):
        states.append(advance_daily_input_operation(states[-1], phase))
    states.append(advance_daily_input_operation(
        states[-1], Phase.REPLACEMENTS_RESOLVED, operator_resolution=resolution,
    ))
    states.append(advance_daily_input_operation(
        states[-1], Phase.FAILED, failure_code="FICTIONAL_AFTER_DECISION",
    ))
    states.append(advance_daily_input_operation(
        states[-1], Phase.ROLLED_BACK, rollback_verified=True,
    ))

    rows = ()
    for state in states:
        rows = append_operation_journal(rows, state)
    reopened = verify_operation_journal_readback(rows, rows)
    assert reopened == tuple(states)
    assert reopened[-1].operator_resolution == resolution
    assert reopened[-1].rollback_verified
    assert not reopened[-1].may_report_success
