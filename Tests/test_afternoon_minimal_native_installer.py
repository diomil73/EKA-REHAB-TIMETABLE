"""Synthetic gates for isolated native PM-only installer, no Excel execution."""
from pathlib import Path

import pytest

from rehab_excel.afternoon_minimal_native_installer import (
    AfternoonMinimalInstallError, _baseline, install_afternoon_minimal_preview,
)
from rehab_excel.afternoon_sheet_preview import AfternoonPreviewError


class FakeUsedRange:
    Address = "$B$1:$P$2"
    Column = 2
    Value2 = (("SYNTHETIC", "VALUE"), (1, 2))
    Formula = (("SYNTHETIC", "VALUE"), (1, 2))

    class _Columns:
        Count = 15

    Columns = _Columns()


class FakeColumn:
    def __init__(self, index):
        self.ColumnWidth = 10.0 + index


class FakeColumns:
    def __call__(self, index):
        return FakeColumn(index)


class FakeSheet:
    UsedRange = FakeUsedRange()
    Columns = FakeColumns()


class FakeSheets:
    def __init__(self):
        self.names = ["MASTER_SCHEDULE", "PATIENT_PLANNER", "PATIENTS"]

    @property
    def Count(self):
        return len(self.names)

    def __call__(self, key):
        if isinstance(key, int):
            return type("NamedSheet", (), {"Name": self.names[key - 1]})()
        if key not in self.names:
            raise KeyError(key)
        return FakeSheet()


class FakeWorkbook:
    Worksheets = FakeSheets()


def test_baseline_captures_all_used_master_widths_and_sheet_formulas():
    before = _baseline(FakeWorkbook())
    assert before[0] == ("MASTER_SCHEDULE", "PATIENT_PLANNER", "PATIENTS")
    assert before[1] == (2, 15, tuple(10.0 + index for index in range(2, 17)))
    assert before[2][0] == "$B$1:$P$2"
    assert before[2][2] == (("SYNTHETIC", "VALUE"), (1, 2))
    assert before[3][2] == (("SYNTHETIC", "VALUE"), (1, 2))
    assert before[4][2] == (("SYNTHETIC", "VALUE"), (1, 2))
    assert len(before) == 5


def test_new_technical_sheet_does_not_change_existing_sheet_invariants():
    book = FakeWorkbook()
    original = _baseline(book)
    book.Worksheets.names.append("AFTERNOON_PHYSIO_SLOTS")
    new = _baseline(book)
    assert new[0] == original[0] + ("AFTERNOON_PHYSIO_SLOTS",)
    assert new[1:] == original[1:]


def test_missing_native_excel_on_other_platform_fails_before_file_write(
    tmp_path, monkeypatch
):
    from rehab_excel import afternoon_minimal_native_installer as module
    monkeypatch.setattr(module.sys, "platform", "linux")
    source = tmp_path / "SYNTHETIC.xlsm"
    source.write_bytes(b"synthetic only")
    output = tmp_path / "UNWRITTEN.xlsm"
    with pytest.raises(AfternoonMinimalInstallError):
        install_afternoon_minimal_preview(source, output)
    assert not output.exists()


def test_only_fresh_distinct_destination_accepted_by_shared_path_gate(tmp_path):
    from rehab_excel.afternoon_sheet_preview import validate_preview_paths
    source = tmp_path / "SYNTH_SOURCE.xlsm"
    source.write_bytes(b"synthetic")
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, source)
    existing = tmp_path / "EXISTING.xlsm"
    existing.write_bytes(b"frozen")
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, existing)


def test_minimal_native_installer_never_calls_broad_patient_ui_or_master_refresh():
    source = (
        Path(__file__).resolve().parents[1] / "Python"
        / "rehab_excel" / "afternoon_minimal_native_installer.py"
    ).read_text(encoding="utf-8")
    assert "install_patient_program_ui(" not in source
    assert "install_patient_program_master_render(" not in source
    assert "_refresh_existing_master_patient_programs(" not in source
    assert "RefreshMasterPatientProgram" not in source
    assert "PATIENT_PLANNER" in source
    assert "_confirm_panel(" in source
    assert 'dst.open("xb")' in source
    assert "book.Save()" in source
    assert "ReadOnly=True" in source
    assert "source changed" in source.lower() or "Source changed" in source


def test_native_installer_checks_source_and_reopen_without_patient_logging():
    source = (
        Path(__file__).resolve().parents[1] / "Python"
        / "rehab_excel" / "afternoon_minimal_native_installer.py"
    ).read_text(encoding="utf-8")
    assert "inspect_afternoon_native_candidate(src)" in source
    assert "patch_afternoon_panel_text(runtime_code, form_code)" in source
    assert "_baseline(readonly)" in source
    assert "_roster(book)" in source
    assert "xlSheetVeryHidden" not in source or "ws.Visible = 2" in source
    assert "item.patient_id" not in source


def test_cli_requires_explicit_disposable_synthetic_confirmation(tmp_path, capsys):
    from tools.build_afternoon_minimal_panel_preview import main
    source = tmp_path / "SYNTH_INPUT.xlsm"
    source.write_bytes(b"synthetic")
    output = tmp_path / "SHOULD_NOT_EXIST.xlsm"
    assert main(["--source", str(source), "--output", str(output)]) == 2
    assert not output.exists()
    assert "SAFETY STOP" in capsys.readouterr().out


def test_existing_vba_components_must_be_fetched_outside_missing_label_guard():
    """A missing runtime or blocked designer must not be treated as label absence."""
    source = (
        Path(__file__).resolve().parents[1] / "Python"
        / "rehab_excel" / "afternoon_minimal_native_installer.py"
    ).read_text(encoding="utf-8")
    component_lookup = source.index('VBComponents("PatientProgramRuntime")', source.index("prior = _baseline(book)"))
    designer_lookup = source.index("designer = form.Designer", component_lookup)
    control_enumeration = source.index("existing_names = tuple(", designer_lookup)
    assert component_lookup < designer_lookup < control_enumeration
    assert "int(designer.Controls.Count)" in source
    assert "if PANEL_LABEL in existing_names:" in source
    assert "existing_label = None" not in source
    assert "Already installed afternoon panel" in source


def test_native_sessions_disable_automatic_workbook_macros():
    source = (
        Path(__file__).resolve().parents[1] / "Python"
        / "rehab_excel" / "afternoon_minimal_native_installer.py"
    ).read_text(encoding="utf-8")
    assert "excel.AutomationSecurity = 3" in source
    assert "verifier.AutomationSecurity = 3" in source
    assert source.index("excel.AutomationSecurity = 3") < source.index(
        'book = excel.Workbooks.Open'
    )
    assert source.index("verifier.AutomationSecurity = 3") < source.index(
        'readonly = verifier.Workbooks.Open'
    )


def test_native_source_readback_tolerates_only_line_endings_and_trailing_blank_lines():
    from rehab_excel.afternoon_minimal_native_installer import _canonical_vba_source
    assert _canonical_vba_source("Option Explicit\r\nSub Test()\r\nEnd Sub\r\n") == (
        _canonical_vba_source("Option Explicit\nSub Test()\nEnd Sub\n\n")
    )
    assert _canonical_vba_source("A\rB") == "A\nB"
    assert _canonical_vba_source("Sub A()\nEnd Sub") != (
        _canonical_vba_source("Sub B()\nEnd Sub")
    )
