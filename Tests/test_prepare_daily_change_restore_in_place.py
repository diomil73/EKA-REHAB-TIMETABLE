import inspect

import pytest

from tools.prepare_daily_change_restore_in_place import (
    SUPPORTED_SOURCE_CHANGE_TYPES,
    _restore_marker_line,
    _source_clear_line,
    prepare_daily_change_restore_payload,
)


def test_source_change_restore_scope_starts_with_patient_absence_and_session_cancellation():
    assert SUPPORTED_SOURCE_CHANGE_TYPES == frozenset(
        {"patient_absence", "session_cancellation"}
    )
    assert "therapist_absence" not in SUPPORTED_SOURCE_CHANGE_TYPES


def test_source_clear_payload_targets_only_the_daily_input_patient_row():
    assert _source_clear_line(23) == "SOURCE_CLEAR\tDAILY_INPUT\t23\tA\tE"


def test_restore_marker_keeps_exact_change_id():
    assert (
        _restore_marker_line("daily-input:2026-10-31:patient:23")
        == "RESTORE_CHANGE\tdaily-input:2026-10-31:patient:23"
    )


def test_restore_preparer_recomputes_truth_with_only_selected_source_row_ignored():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "current_daily.source_changes" in source
    assert "len(matches) != 1" in source
    assert "ignore_rows=(change.source_row,)" in source
    assert "allow_empty=True" in source
    assert "restored_daily.absences" in source
    assert "restored_daily.cancellations" in source


def test_restore_preparer_preserves_active_replacements_and_rebuilds_both_views():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "read_daily_replacement_facts(" in source
    assert "replacements=replacements" in source
    assert "build_patient_centric_preview_plan(" in source
    assert "compose_outpatient_daily_plan(" in source
    assert "build_master_daily_overlay_plan(" in source
    assert "compose_master_daily_overlay(" in source
    assert "validate_write_plan(plan)" in source


def test_restore_preparer_is_read_only_and_fail_closed_for_therapist_absence():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "file_sha256(source)" in source
    assert "Workbook changed during daily-change restore validation" in source
    assert "coordinated rollback of related replacement facts" in source
    assert "Workbooks.Open" not in source
    assert "workbook.Save" not in source
    assert "shutil.copy" not in source


def test_restore_payload_orders_source_clear_and_restore_marker_before_projection_patches():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "_source_clear_line(change.source_row)" in source
    assert "_restore_marker_line(change.change_id)" in source
    assert "*(_patch_line(patch) for patch in plan.patches)" in source


def test_empty_change_id_is_rejected_before_any_restore_work(tmp_path):
    with pytest.raises(Exception, match="Workbook not found"):
        prepare_daily_change_restore_payload(
            tmp_path / "missing.xlsm",
            tmp_path / "payload.tsv",
            change_id="",
        )
