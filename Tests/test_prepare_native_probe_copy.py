"""Synthetic private XLSM copy safety invariants."""
from pathlib import Path
from zipfile import ZipFile

import pytest

from tools.prepare_native_probe_copy import ProbeCopyError, create_probe_copy


def source_xlsm(tmp_path):
    source = tmp_path / "source.xlsm"
    with ZipFile(source, "w") as z:
        z.writestr("xl/vbaProject.bin", b"SYNTHETIC")
        z.writestr("xl/workbook.xml", b"<workbook/>")
    return source


def test_copy_is_exact_and_source_unchanged(tmp_path):
    source = source_xlsm(tmp_path)
    output = tmp_path / "probe.xlsm"
    before = source.read_bytes()
    outcome = create_probe_copy(source, output)
    assert outcome["byte_exact"] and outcome["source_unchanged"]
    assert output.read_bytes() == before == source.read_bytes()


def test_existing_destination_is_never_overwritten(tmp_path):
    source = source_xlsm(tmp_path)
    output = tmp_path / "probe.xlsm"
    output.write_bytes(b"EXTERNAL FILE")
    with pytest.raises(ProbeCopyError, match="refusing overwrite"):
        create_probe_copy(source, output)
    assert output.read_bytes() == b"EXTERNAL FILE"


def test_reject_source_equals_output(tmp_path):
    source = source_xlsm(tmp_path)
    with pytest.raises(ProbeCopyError, match="Distinct"):
        create_probe_copy(source, source)


def test_reject_missing_directory_without_creating_anything(tmp_path):
    source = source_xlsm(tmp_path)
    output = tmp_path / "not-created" / "probe.xlsm"
    with pytest.raises(ProbeCopyError, match="parent folder"):
        create_probe_copy(source, output)
    assert not output.parent.exists()


def test_exclusive_create_race_does_not_delete_other_file(tmp_path, monkeypatch):
    source = source_xlsm(tmp_path)
    output = tmp_path / "probe.xlsm"
    native_open = Path.open

    def race(path, mode="r", *args, **kwargs):
        if Path(path) == output and mode == "xb":
            output.write_bytes(b"OTHER PROCESS WON")
            raise FileExistsError("already exists")
        return native_open(path, mode, *args, **kwargs)

    monkeypatch.setattr(Path, "open", race)
    with pytest.raises(FileExistsError):
        create_probe_copy(source, output)
    assert output.read_bytes() == b"OTHER PROCESS WON"


def test_corrupted_vba_source_does_not_produce_output(tmp_path):
    source = tmp_path / "source.xlsm"
    with ZipFile(source, "w") as z:
        z.writestr("xl/workbook.xml", "SYNTHETIC")
    output = tmp_path / "probe.xlsm"
    with pytest.raises(ProbeCopyError, match="VBA"):
        create_probe_copy(source, output)
    assert not output.exists()
