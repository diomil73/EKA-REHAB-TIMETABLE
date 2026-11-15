import inspect

from tools import audit_vba_project_inventory as module


def test_component_type_labels_are_stable_and_non_sensitive():
    assert module.component_type_label(1) == "STANDARD_MODULE"
    assert module.component_type_label(2) == "CLASS_MODULE"
    assert module.component_type_label(3) == "USERFORM"
    assert module.component_type_label(100) == "DOCUMENT_MODULE"
    assert module.component_type_label(77) == "OTHER_TYPE_77"


def test_native_inventory_is_read_only_and_disables_macros_and_events():
    source = inspect.getsource(module.audit_one)
    assert "AutomationSecurity = 3" in source
    assert "EnableEvents = False" in source
    assert "Calculation = -4135" in source
    assert "manual calculation mode unavailable" in source
    assert "calculation_manual" in source
    assert "ReadOnly=True" in source
    assert "UpdateLinks=0" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert "workbook.Save(" not in source
    assert "workbook.Save()" not in source
    assert ".Lines(" not in source
    assert ".Caption" not in source
    assert "SOURCE HASH UNCHANGED" in source


def test_duplicate_workbook_bytes_are_audited_only_once(tmp_path, monkeypatch, capsys):
    first = tmp_path / "first.xlsm"
    duplicate = tmp_path / "duplicate.xlsm"
    other = tmp_path / "other.xlsm"
    first.write_bytes(b"same workbook bytes")
    duplicate.write_bytes(b"same workbook bytes")
    other.write_bytes(b"different workbook bytes")

    calls = []

    def fake_audit(path):
        calls.append(path.name)
        return 0, module.file_sha256(path)

    monkeypatch.setattr(module, "audit_one", fake_audit)

    assert module.audit_many([first, duplicate, other]) == 0
    assert calls == ["first.xlsm", "other.xlsm"]

    output = capsys.readouterr().out
    assert "DUPLICATE BYTE-IDENTICAL CONTENT" in output
    assert "ALL DISTINCT VBA PROJECT INVENTORIES COMPLETE" in output


def test_duplicate_audit_does_not_hide_a_failed_distinct_candidate(tmp_path, monkeypatch):
    first = tmp_path / "first.xlsm"
    duplicate = tmp_path / "duplicate.xlsm"
    other = tmp_path / "other.xlsm"
    first.write_bytes(b"same workbook bytes")
    duplicate.write_bytes(b"same workbook bytes")
    other.write_bytes(b"other bytes")

    def fake_audit(path):
        if path.name == "other.xlsm":
            return 2, None
        return 0, module.file_sha256(path)

    monkeypatch.setattr(module, "audit_one", fake_audit)

    assert module.audit_many([first, duplicate, other]) == 2



def test_duplicate_failed_audit_status_is_preserved(tmp_path, monkeypatch, capsys):
    first = tmp_path / "first.xlsm"
    duplicate = tmp_path / "duplicate.xlsm"
    first.write_bytes(b"same workbook bytes")
    duplicate.write_bytes(b"same workbook bytes")
    calls = []

    def failed_audit(path):
        calls.append(path.name)
        return 2, None

    monkeypatch.setattr(module, "audit_one", failed_audit)

    assert module.audit_many([first, duplicate]) == 2
    assert calls == ["first.xlsm"]
    assert "Original audit status: 2" in capsys.readouterr().out

def test_native_inventory_reports_com_failure_stage_and_hresult_without_source():
    source = inspect.getsource(module.audit_one)
    assert 'stage = "starting isolated Excel instance"' in source
    assert 'stage = "accessing VBA project object model"' in source
    assert 'stage = "enumerating VBA components"' in source
    assert "COM HRESULT:" in source
    assert "COM DESCRIPTION:" in source
    assert "CodeModule.Lines(" not in source
    assert ".Formula" not in source
    assert "workbook.Save(" not in source
