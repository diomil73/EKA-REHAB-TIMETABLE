import inspect

import pytest

from tools.prepare_daily_change_restore_in_place import (
    SUPPORTED_SOURCE_CHANGE_TYPES,
    _build_clean_day_restore_plan,
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


def test_restore_preparer_recomputes_truth_with_all_selected_source_rows_ignored_atomically():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "current_daily.source_changes" in source
    assert "missing = tuple" in source
    assert "ignore_rows = tuple(sorted({change.source_row for change in changes}))" in source
    assert "ignore_rows=ignore_rows" in source
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


def test_restore_to_fully_active_day_uses_bounded_clean_day_path():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "non_active_after_restore" in source
    assert "if non_active_after_restore:" in source
    assert "current_states = build_daily_session_states(" in source
    assert "plan = _build_clean_day_restore_plan(" in source


def test_clean_day_restore_rebuilds_only_touched_operational_cells_from_base():
    source = inspect.getsource(_build_clean_day_restore_plan)
    assert "rebuild_multi_member_groups=False" in source
    assert "current_preview.bindings" in source
    assert "render_base_slot(" in source
    assert 'sheet="THERAPIST_DAILY"' in source
    assert 'source_tag="daily_change_restore_base_slot"' in source
    assert "_base_slot_runs(render)" in source


def test_clean_day_restore_resets_master_treatment_and_status_presentation():
    source = inspect.getsource(_build_clean_day_restore_plan)
    assert "build_master_daily_overlay_plan(" in source
    assert "build_master_projection(" in source
    assert 'patch.source_tag == "master_daily_non_active_overlay"' in source
    assert 'source_tag="master_daily_restore_overlay"' in source
    assert 'patch.source_tag == "master_daily_patient_status_overlay"' in source
    assert 'source_tag="master_daily_restore_status"' in source
    assert "_plain_runs(value)" in source


def test_clean_day_restore_does_not_weaken_generic_preview_empty_state_guard():
    source = inspect.getsource(_build_clean_day_restore_plan)
    assert "current_states" in source
    assert "build_patient_centric_preview_plan(" in source
    assert "rebuild_multi_member_groups=False" in source
    preparer = inspect.getsource(prepare_daily_change_restore_payload)
    assert "states=states" in preparer
    assert "non_active_after_restore" in preparer


def test_restore_preparer_is_read_only_and_fail_closed_for_therapist_absence():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "file_sha256(source)" in source
    assert "Workbook changed during daily-change restore validation" in source
    assert "coordinated rollback of related replacement facts" in source
    assert "Workbooks.Open" not in source
    assert "workbook.Save" not in source
    assert "shutil.copy" not in source


def test_restore_payload_orders_all_source_clears_and_markers_before_projection_patches():
    source = inspect.getsource(prepare_daily_change_restore_payload)
    assert "*(_source_clear_line(change.source_row) for change in changes)" in source
    assert "*(_restore_marker_line(change.change_id) for change in changes)" in source
    assert "*(_patch_line(patch) for patch in plan.patches)" in source


def test_restore_cli_accepts_repeated_change_ids():
    source = inspect.getsource(__import__("tools.prepare_daily_change_restore_in_place", fromlist=["main"]).main)
    assert 'parser.add_argument("--change-id", action="append", default=[])' in source
    assert "change_ids=tuple(args.change_id)" in source


def test_missing_workbook_fails_before_restore_mutation(tmp_path):
    with pytest.raises(Exception, match="Workbook not found"):
        prepare_daily_change_restore_payload(
            tmp_path / "missing.xlsm",
            tmp_path / "payload.tsv",
            change_ids=("daily-input:2026-10-31:patient:23",),
        )
