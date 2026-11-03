import inspect

from tools import prepare_close_day_in_place as module


def test_close_day_preparation_is_read_only_for_workbook():
    source = inspect.getsource(module.prepare_close_day_payload)

    assert "file_sha256(source)" in source
    assert "assert_application_contract(source)" in source
    assert "read_finalized_day_snapshots" in source
    assert "read_patient_session_history_rows" in source
    assert "read_therapist_day_history_rows" in source
    assert "read_day_history_events" in source
    assert "build_finalized_day_snapshot" in source
    assert "build_patient_session_history_rows" in source
    assert "build_therapist_day_history_rows" in source
    assert "build_day_finalized_event" in source

    assert "append_finalized_day_snapshot" not in source
    assert "append_patient_session_history_rows" not in source
    assert "append_therapist_day_history_rows" not in source
    assert "append_day_history_events" not in source
    assert "shutil.copy" not in source


def test_close_day_preparation_allows_clean_day_without_daily_input_changes():
    source = inspect.getsource(module.prepare_close_day_payload)

    assert "allow_empty=True" in source


def test_close_day_payload_contains_all_atomic_blocks_including_patient_history():
    source = inspect.getsource(module.prepare_close_day_payload)

    assert '"SNAPSHOT"' in source
    assert '_line("PATIENT", patient_values_for_payload(row))' in source
    assert '"THERAPIST"' in source
    assert '"EVENT"' in source
    assert '"patient_history_rows": len(patient_rows)' in source
    assert '"payload_path": str(payload)' in source
    assert '"source_unchanged": True' in source


def test_close_day_patient_history_is_bound_to_v1_snapshot():
    source = inspect.getsource(module.prepare_close_day_payload)

    assert "snapshot_id=snapshot.snapshot_id" in source
    assert "snapshot_version=snapshot.version" in source
    assert "Patient/session history already exists for snapshot" in source


def test_close_day_payload_escape_round_trip_contract():
    escaped = module._escape("A\\B\tC\nD\rE")
    assert escaped == "A\\\\B\\tC\\nD\\rE"
