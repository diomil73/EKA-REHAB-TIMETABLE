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
