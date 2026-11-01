import inspect

from tools import inspect_day_amendment_outcome as module


def test_amendment_outcome_inspector_requires_contiguous_version_chain():
    source = inspect.getsource(module.inspect_day_amendment_outcome)

    assert "FinalizedSnapshotStatus.CURRENT_FINAL" in source
    assert "FinalizedSnapshotStatus.SUPERSEDED" in source
    assert "current.supersedes_snapshot_id" in source
    assert "current.version != previous.version + 1" in source
    assert "current.unresolved_sessions != 0" in source


def test_amendment_outcome_inspector_requires_linked_history_and_events():
    source = inspect.getsource(module.inspect_day_amendment_outcome)

    assert "read_therapist_day_history_rows" in source
    assert "DAY_AMENDMENT_STARTED" in source
    assert "DAY_AMENDMENT_FINALIZED" in source
    assert "related_event_id" in source
    assert "related_snapshot_id" in source


def test_amendment_outcome_inspector_has_stable_success_and_failure_markers():
    main_source = inspect.getsource(module.main)

    assert "DAY AMENDMENT OUTCOME OK" in main_source
    assert "DAY AMENDMENT INSPECTION FAILED" in main_source
