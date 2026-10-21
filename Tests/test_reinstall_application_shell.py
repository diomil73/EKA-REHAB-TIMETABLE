from __future__ import annotations

import inspect

from tools.reinstall_application_shell import reinstall_application_shell


def test_shell_repair_keeps_source_immutable_and_uses_fresh_output_copy():
    source = inspect.getsource(reinstall_application_shell)

    assert "if source == output" in source
    assert "shutil.copy2(source, output)" in source
    assert 'Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)' in source
    assert "install_application_shell(workbook.VBProject, workbook=workbook)" in source
    assert source.count("workbook.Save()") == 1


def test_shell_repair_requires_xlsm_vba_and_windows_excel():
    source = inspect.getsource(reinstall_application_shell)

    assert 'source.suffix.casefold() != ".xlsm"' in source
    assert "if not _has_vba(source):" in source
    assert 'sys.platform != "win32"' in source
    assert 'DispatchEx("Excel.Application")' in source
    assert "if not _has_vba(output):" in source
