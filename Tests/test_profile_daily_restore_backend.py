from __future__ import annotations

import inspect

from tools.profile_daily_restore_backend import (
    _profiled_apply_factory,
    profile_restore_backend,
)


def test_profiled_apply_captures_native_excel_subphases():
    source = inspect.getsource(_profiled_apply_factory)

    for marker in (
        'timings["excel_dispatch"]',
        'timings["workbook_open"]',
        'timings["master_projection"]',
        'timings["daily_patches"]',
        'timings["statistics"]',
        'timings["replacement_fact_rows"]',
        'timings["ledger_rows"]',
        'timings["ledger_refresh"]',
        'timings["workbook_save"]',
        'timings["workbook_close"]',
        'timings["excel_quit"]',
        'timings["excel_output_total"]',
    ):
        assert marker in source


def test_profiler_restores_production_apply_method_after_run():
    source = inspect.getsource(profile_restore_backend)

    assert "original_apply = restore_tool._RestoreOutputBackend.apply" in source
    assert "restore_tool._RestoreOutputBackend.apply = _profiled_apply_factory(timings)" in source
    assert "restore_tool._RestoreOutputBackend.apply = original_apply" in source
    assert 'timings["backend_total"]' in source
    assert 'timings["pre_excel_and_verification"]' in source
    assert 'timings["pre_excel_unattributed"]' in source


def test_profiler_captures_pre_excel_pipeline_subphases():
    source = inspect.getsource(profile_restore_backend)

    for marker in (
        'patch(restore_tool.shutil, "copy2", "stage_copy")',
        'patch(restore_tool, "_sha256", "sha256_checks")',
        'patch(restore_tool, "read_daily_replacement_fact_rows", "read_replacement_fact_rows")',
        'patch(restore_tool, "validate_batch_restore_request", "validate_restore_ids")',
        'patch(restore_tool, "read_patient_registry", "read_patients")',
        'patch(restore_tool, "read_unified_base_schedule", "read_base_schedule")',
        'patch(restore_tool, "read_daily_input", "read_daily_input")',
        'patch(restore_tool, "_active_replacements_after_restore", "derive_replacement_facts")',
        'patch(restore_tool, "build_master_projection", "build_master_projection")',
        'patch(restore_tool, "build_patient_centric_preview_plan", "build_patient_preview")',
        'patch(restore_tool, "build_master_daily_overlay_plan", "build_master_overlay")',
        'patch(restore_tool, "apply_write_plan_to_copy", "apply_write_plan_total")',
        'patch(restore_tool, "_has_vba", "vba_checks")',
        'patch(native_excel, "validate_write_plan", "native_write_plan_validation")',
        'patch(native_excel, "_sha256", "native_write_hashes")',
    ):
        assert marker in source

    assert 'restore_fact_batch' not in source
    assert "for owner, name, original in reversed(originals):" in source
    assert "setattr(owner, name, original)" in source
