from __future__ import annotations

import inspect
from pathlib import Path

from tools.build_outpatient_recurring_smoke import (
    _install_conflict_guard_with_lock_retry,
    build_outpatient_recurring_smoke,
)


def test_smoke_builder_installs_create_edit_then_conflict_guard():
    source = inspect.getsource(build_outpatient_recurring_smoke)

    create_pos = source.index("create_outpatient_schedule_form_preview(")
    edit_pos = source.index("install_outpatient_schedule_edit_support_in_workbook(stage)")
    conflict_pos = source.index("_install_conflict_guard_with_lock_retry(")

    assert create_pos < edit_pos < conflict_pos
    assert 'stage = output.with_name(f".{output.stem}.outpatient-form-stage{output.suffix}")' in source
    assert "finally:" in source
    assert "_safe_unlink(stage)" in source


def test_smoke_builder_never_targets_source_for_mutation():
    source = inspect.getsource(build_outpatient_recurring_smoke)

    assert "create_outpatient_schedule_form_preview(source, stage" in source
    assert "install_outpatient_schedule_edit_support_in_workbook(stage)" in source
    assert "_install_conflict_guard_with_lock_retry(" in source
    assert "install_outpatient_schedule_edit_support_in_workbook(source)" not in source
    assert "create_outpatient_conflict_guard_preview(source, output" not in source
    assert "source_before = _sha256(source)" in source
    assert "_sha256(source) != source_before" in source


def test_conflict_guard_retry_handles_only_winerror_32():
    source = inspect.getsource(_install_conflict_guard_with_lock_retry)

    assert 'getattr(exc, "winerror", None) != 32' in source
    assert "time.sleep(delay_seconds)" in source
    assert "attempts: int = 20" in source
    assert "delay_seconds: float = 0.25" in source


def test_smoke_builder_requires_distinct_xlsm_paths(tmp_path: Path):
    source = tmp_path / "source.xlsm"
    source.write_bytes(b"placeholder")

    try:
        build_outpatient_recurring_smoke(source, source, overwrite=True)
    except RuntimeError as exc:
        assert "Output must differ from source" in str(exc)
    else:
        raise AssertionError("same source/output path must be rejected")
