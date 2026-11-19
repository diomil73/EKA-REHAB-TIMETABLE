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
