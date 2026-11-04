import inspect

from tools import audit_vba_write_surfaces as module


def test_vba_write_surface_scan_finds_sheet_references_and_mutations():
    lines = [
        "Public Sub SavePatient()",
        'Set ws = ThisWorkbook.Worksheets("PATIENTS")',
        "ws.Cells(rowIndex, 3).Value = displayName",
        "ws.Unprotect",
        "ws.Protect",
        'ws.Range("A1").ClearContents',
        "End Sub",
    ]
    matches = module._matches(lines)
    assert [item[1] for item in matches] == ["SavePatient"] * 5
    assert any("Worksheets" in item[2] for item in matches)
    assert any(".Value" in item[2] for item in matches)
    assert any(".Unprotect" in item[2] for item in matches)
    assert any(".Protect" in item[2] for item in matches)


def test_vba_write_surface_scan_tracks_procedures_and_line_numbers():
    matches = module._matches([
        "Option Explicit",
        "Public Sub FirstAction()",
        "ws.Range(\"A1\").ClearContents",
        "End Sub",
        "Private Function SecondAction() As Boolean",
        "ws.Cells(2, 1).Formula = \"=1+1\"",
        "End Function",
    ])
    assert matches == [
        (3, "FirstAction", 'ws.Range("A1").ClearContents'),
        (6, "SecondAction", 'ws.Cells(2, 1).Formula = "=1+1"'),
    ]


def test_vba_source_audit_opens_read_only_and_never_saves_or_mutates():
    source = inspect.getsource(module.audit)
    assert "ReadOnly=True" in source
    assert "UpdateLinks=0" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert "workbook.Save(" not in source
    assert "workbook.Save()" not in source
    assert "ws.Unprotect(" not in source
    assert "ws.Protect(" not in source
