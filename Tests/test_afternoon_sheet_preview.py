"""Synthetic-only safety checks for native afternoon technical-sheet preview."""
from pathlib import Path

import pytest

from rehab_excel.afternoon_sheet_preview import (
    AfternoonPreviewError, validate_preview_paths,
)


def test_distinct_xlsm_paths_are_eligible_for_future_native_preview(tmp_path):
    source = tmp_path / "SYNTH_DIOMIL_SOURCE.xlsm"
    source.write_bytes(b"synthetic-placeholder-not-a-native-workbook")
    output = tmp_path / "SYNTH_DIOMIL_AFTERNOON.xlsm"
    assert validate_preview_paths(source, output) == (source.resolve(), output.resolve())
    assert not output.exists()


def test_refuses_same_path_even_when_source_is_present(tmp_path):
    source = tmp_path / "SYNTH.xlsm"
    source.write_bytes(b"synthetic")
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, source)


def test_refuses_overwrite_of_existing_preview(tmp_path):
    source = tmp_path / "SYNTH_SOURCE.xlsm"
    output = tmp_path / "SYNTH_OUTPUT.xlsm"
    source.write_bytes(b"synthetic-source")
    output.write_bytes(b"synthetic-frozen-evidence")
    before = output.read_bytes()
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, output)
    assert output.read_bytes() == before


@pytest.mark.parametrize("extension", [".xlsx", ".xls", ".txt"])
def test_refuses_non_macro_preview_output(tmp_path, extension):
    source = tmp_path / "SYNTH_SOURCE.xlsm"
    source.write_bytes(b"synthetic")
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, tmp_path / ("SYNTH_OUTPUT" + extension))


def test_refuses_non_macro_source(tmp_path):
    source = tmp_path / "SYNTH_SOURCE.xlsx"
    source.write_bytes(b"synthetic")
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, tmp_path / "OUTPUT.xlsm")


def test_refuses_missing_output_folder(tmp_path):
    source = tmp_path / "SYNTH_SOURCE.xlsm"
    source.write_bytes(b"synthetic")
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, tmp_path / "missing" / "OUTPUT.xlsm")


def test_refuses_symlink_output(tmp_path):
    source = tmp_path / "SYNTH_SOURCE.xlsm"
    source.write_bytes(b"synthetic")
    existing = tmp_path / "EXISTING.xlsm"
    existing.write_bytes(b"synthetic-frozen")
    link = tmp_path / "OUTPUT.xlsm"
    try:
        link.symlink_to(existing)
    except (OSError, NotImplementedError):
        pytest.skip("Symlinks unavailable")
    with pytest.raises(AfternoonPreviewError):
        validate_preview_paths(source, link)


def test_native_builder_does_not_run_without_windows(tmp_path, monkeypatch):
    from rehab_excel import afternoon_sheet_preview
    source = tmp_path / "SYNTH_SOURCE.xlsm"
    source.write_bytes(b"synthetic")
    output = tmp_path / "OUTPUT.xlsm"
    monkeypatch.setattr(afternoon_sheet_preview.sys, "platform", "linux")
    with pytest.raises(AfternoonPreviewError):
        afternoon_sheet_preview.build_afternoon_sheet_preview(source, output)
    assert not output.exists()
