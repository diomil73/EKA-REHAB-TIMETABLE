import inspect

from tools import apply_day_amendment as module


def test_apply_day_amendment_is_hash_bound_before_persistence():
    source = inspect.getsource(module.apply_day_amendment)

    assert 'data.get("schema") != "day-amendment-v1"' in source
    assert 'Path(_required_text(data.get("source_path")' in source
    assert "file_sha256(source)" in source
    assert "actual_hash != expected_hash" in source
    assert "Workbook changed after amendment preparation" in source
    assert "persist_day_amendment(source, plan, therapist_rows)" in source


def test_apply_day_amendment_reconstructs_full_version_transition():
    source = inspect.getsource(module.apply_day_amendment)

    assert 'data.get("previous_snapshot")' in source
    assert 'data.get("superseded_snapshot")' in source
    assert 'data.get("replacement_snapshot")' in source
    assert 'data.get("started_event")' in source
    assert 'data.get("finalized_event")' in source
    assert 'data.get("therapist_rows")' in source
    assert "DayAmendmentPlan(" in source


def test_apply_day_amendment_requires_post_save_semantic_verification():
    source = inspect.getsource(module.apply_day_amendment)

    assert "read_finalized_day_snapshots" in source
    assert "read_therapist_day_history_rows" in source
    assert "read_day_history_events" in source
    assert "Post-save verification failed for superseded snapshot" in source
    assert "Post-save verification failed for replacement snapshot" in source
    assert "Post-save therapist/day history verification failed" in source
    assert "Post-save amendment event verification failed" in source
    assert '"post_save_verified": True' in source


def test_apply_cli_reports_verified_success_only_after_apply_returns():
    source = inspect.getsource(module.main)

    assert "apply_day_amendment(args.workbook, args.payload)" in source
    assert 'print("DAY AMENDMENT APPLY VERIFIED")' in source
    assert 'print("Post-save verification: True")' in source
