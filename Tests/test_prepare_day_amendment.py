import inspect

from tools import prepare_day_amendment as module


def test_day_amendment_preparation_is_read_only_and_hash_bound():
    source = inspect.getsource(module.prepare_day_amendment)

    assert "file_sha256(source)" in source
    assert "assert_application_contract(source)" in source
    assert "read_finalized_day_snapshots" in source
    assert "build_day_amendment_plan" in source
    assert "build_therapist_day_history_rows" in source
    assert '"source_sha256": source_hash' in source
    assert '"source_unchanged": True' in source

    assert "persist_day_amendment" not in source
    assert "append_finalized_day_snapshot" not in source
    assert "append_therapist_day_history_rows" not in source
    assert "append_day_history_events" not in source
    assert "workbook.Save" not in source


def test_day_amendment_preparation_requires_exactly_one_current_snapshot():
    source = inspect.getsource(module.prepare_day_amendment)

    assert "current_only=True" in source
    assert "len(current) != 1" in source
    assert "Expected exactly one CURRENT_FINAL snapshot" in source


def test_day_amendment_payload_carries_full_version_transition():
    source = inspect.getsource(module.prepare_day_amendment)

    assert '"schema": "day-amendment-v1"' in source
    assert '"previous_snapshot"' in source
    assert '"superseded_snapshot"' in source
    assert '"replacement_snapshot"' in source
    assert '"therapist_rows"' in source
    assert '"started_event"' in source
    assert '"finalized_event"' in source
    assert "previous.version + 1" in source
    assert "unresolved_sessions" in source


def test_day_amendment_preparation_rejects_event_and_history_collisions():
    source = inspect.getsource(module.prepare_day_amendment)

    assert "read_therapist_day_history_rows" in source
    assert "read_day_history_events" in source
    assert "Replacement therapist/day history already exists" in source
    assert "Amendment EventId already exists" in source
