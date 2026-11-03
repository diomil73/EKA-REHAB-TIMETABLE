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


def test_amendment_outcome_inspector_requires_patient_history_for_both_versions():
    source = inspect.getsource(module.inspect_day_amendment_outcome)

    assert "read_patient_session_history_rows" in source
    assert "read_current_patient_session_history_rows" in source
    assert "previous.snapshot_id" in source
    assert "current.snapshot_id" in source
    assert "No patient/session history rows for superseded snapshot" in source
    assert "No patient/session history rows for current snapshot" in source
    assert "Operational patient history is not restricted to CURRENT_FINAL" in source
    assert "Operational patient history does not match current snapshot rows" in source


def test_amendment_outcome_inspector_prints_patient_history_counts():
    main_source = inspect.getsource(module.main)

    assert "Previous patient/session history rows" in main_source
    assert "Current patient/session history rows" in main_source
    assert "Operational CURRENT_FINAL patient rows" in main_source


def test_amendment_outcome_inspector_has_stable_success_and_failure_markers():
    main_source = inspect.getsource(module.main)

    assert "DAY AMENDMENT OUTCOME OK" in main_source
    assert "DAY AMENDMENT INSPECTION FAILED" in main_source
