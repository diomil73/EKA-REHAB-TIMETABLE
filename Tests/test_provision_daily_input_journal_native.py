from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile

import pytest

from Python.tools.provision_daily_input_journal_native import (
    NativeJournalProvisionError,
    _vba_sha256,
    provision_daily_input_journal,
)


def make_xlsm(path: Path):
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"synthetic-vba-project")
        archive.writestr("[Content_Types].xml", b"<Types/>")


def test_vba_digest_reads_exact_project_stream(tmp_path):
    path = tmp_path / "source.xlsm"
    make_xlsm(path)
    assert _vba_sha256(path) == sha256(b"synthetic-vba-project").hexdigest()


def test_native_failure_removes_only_its_disposable_copy_and_preserves_source(tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "probe.xlsm"
    make_xlsm(source)
    before = source.read_bytes()

    def fail_to_start_excel():
        raise RuntimeError("synthetic Excel startup failure")

    with pytest.raises(
        NativeJournalProvisionError,
        match=r"Stage=start isolated Excel instance; cause=RuntimeError; detail=synthetic Excel startup failure",
    ):
        provision_daily_input_journal(source, output, excel_factory=fail_to_start_excel)

    assert source.read_bytes() == before
    assert not output.exists()


def test_existing_output_is_never_overwritten(tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "probe.xlsm"
    make_xlsm(source)
    output.write_bytes(b"operator-owned")
    with pytest.raises(NativeJournalProvisionError, match="refusing overwrite"):
        provision_daily_input_journal(source, output, excel_factory=lambda: None)
    assert output.read_bytes() == b"operator-owned"
    
def test_failure_diagnostic_redacts_local_source_and_output_paths(tmp_path):
    source = tmp_path / "private-source.xlsm"
    output = tmp_path / "private-output.xlsm"
    make_xlsm(source)

    def fail_to_start_excel():
        raise RuntimeError(f"cannot start for {source} and {output}")

    with pytest.raises(NativeJournalProvisionError) as caught:
        provision_daily_input_journal(source, output, excel_factory=fail_to_start_excel)
    message = str(caught.value)
    assert "Stage=start isolated Excel instance" in message
    assert "<SOURCE>" in message
    assert "<OUTPUT>" in message
    assert str(source) not in message
    assert str(output) not in message
    assert not output.exists()
