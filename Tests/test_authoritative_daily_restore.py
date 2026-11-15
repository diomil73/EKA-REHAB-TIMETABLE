import inspect

from tools.restore_daily_replacement_fact import restore_replacement_and_recompute


def test_restore_pipeline_removes_fact_before_reconstructing_daily_state():
    source = inspect.getsource(restore_replacement_and_recompute)

    restore_pos = source.index("restore_daily_replacement_fact(stage, replacement_id)")
    read_pos = source.index("read_daily_replacement_facts(")
    state_pos = source.index("build_daily_session_states(")

    assert restore_pos < read_pos < state_pos


def test_restore_pipeline_rebuilds_master_recurring_before_reapplying_daily_overlay():
    source = inspect.getsource(restore_replacement_and_recompute)

    recurring_pos = source.index("refresh_master_projection_in_place(stage)")
    preview_pos = source.index("build_patient_centric_preview_plan(")
    overlay_pos = source.index("build_master_daily_overlay_plan(")

    assert recurring_pos < preview_pos < overlay_pos


def test_restore_pipeline_rewrites_statistics_and_ledger_from_recomputed_truth():
    source = inspect.getsource(restore_replacement_and_recompute)

    assert "write_daily_statistics(output, target_date, statistics)" in source
    assert "_mark_ledger_restored(output, replacement_id)" in source
    assert "Cells(ledgerRow, 13).Copy" not in source
    assert "snapshot" not in source.casefold()


def test_restore_pipeline_preserves_source_and_vba():
    source = inspect.getsource(restore_replacement_and_recompute)

    assert "source_before = _sha256(source)" in source
    assert "source_after = _sha256(source)" in source
    assert "_has_vba(source) == _has_vba(output)" in source
