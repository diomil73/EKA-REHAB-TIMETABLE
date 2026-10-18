from __future__ import annotations

import inspect
from pathlib import Path

from tools.build_outpatient_recurring_smoke import build_outpatient_recurring_smoke


def test_smoke_builder_uses_one_excel_session_for_all_vba_installation():
    source = inspect.getsource(build_outpatient_recurring_smoke)

    assert 'DispatchEx("Excel.Application")' in source
    assert source.count("Workbooks.Open(") == 1
    assert "_install_schedule_form(vbproject)" in source
    assert "install_outpatient_schedule_edit_support(vbproject)" in source
    assert "install_outpatient_conflict_guard(vbproject)" in source
    assert "create_outpatient_schedule_form_preview" not in source
    assert "create_outpatient_conflict_guard_preview" not in source
    assert ".outpatient-form-stage" not in source.casefold()
    assert "stage =" not in source.casefold()


def test_smoke_builder_never_targets_source_for_mutation():
    source = inspect.getsource(build_outpatient_recurring_smoke)

    assert "source_before = _sha256(source)" in source
    assert "shutil.copy2(source, output)" in source
    assert "Workbooks.Open(\n            str(output)" in source
    assert "Workbooks.Open(str(source)" not in source
    assert "_sha256(source) != source_before" in source


def test_smoke_builder_saves_only_after_all_installers_and_verification():
    source = inspect.getsource(build_outpatient_recurring_smoke)

    form_pos = source.index("_install_schedule_form(vbproject)")
    edit_pos = source.index("install_outpatient_schedule_edit_support(vbproject)")
    conflict_pos = source.index("install_outpatient_conflict_guard(vbproject)")
    verify_pos = source.index("_verify_components(vbproject)")
    save_pos = source.index("workbook.Save()")

    assert form_pos < edit_pos < conflict_pos < verify_pos < save_pos


def test_smoke_builder_reports_exact_native_install_phase_on_failure():
    source = inspect.getsource(build_outpatient_recurring_smoke)

    for label in (
        "Excel startup",
        "workbook open",
        "schedule form install",
        "edit support install",
        "conflict guard install",
        "component verification",
        "workbook save",
    ):
        assert f'phase = "{label}"' in source
    assert "failed during {failed_phase}" in source


def test_smoke_builder_requires_distinct_xlsm_paths(tmp_path: Path):
    source = tmp_path / "source.xlsm"
    source.write_bytes(b"placeholder")

    try:
        build_outpatient_recurring_smoke(source, source, overwrite=True)
    except RuntimeError as exc:
        assert "Output must differ from source" in str(exc)
    else:
        raise AssertionError("same source/output path must be rejected")
