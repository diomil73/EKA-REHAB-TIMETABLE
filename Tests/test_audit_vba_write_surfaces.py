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



def test_write_surface_operation_classification_is_sanitized():
    assert module._operation('Set ws = ThisWorkbook.Worksheets("PATIENTS")') == (
        "worksheet_reference", "PATIENTS"
    )
    assert module._operation('ws.Cells(2, 1).Value = patientName') == (
        "cell_value_or_formula_write", None
    )
    assert module._operation("ws.Unprotect") == ("protection_unprotect", None)
    assert module._operation("ws.Protect") == ("protection_protect", None)
    assert module._operation("ThisWorkbook.Save") == ("workbook_save", None)
    assert module._operation('ws.Visible = xlSheetVeryHidden') == (
        "visibility_navigation_or_action_binding", None
    )


def test_write_surface_audit_disables_macros_and_verifies_hash():
    source = inspect.getsource(module.audit)
    assert "AutomationSecurity = 3" in source
    assert "EnableEvents = False" in source
    assert "ReadOnly=True" in source
    assert "UpdateLinks=0" in source
    assert "AddToMru=False" in source
    assert "_fingerprint(path)" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert "operation={operation}" in source
    assert "source}" not in source



def test_write_surface_audit_escapes_unicode_component_and_sheet_names():
    assert module._safe_label("Φύλλο1").startswith(chr(92) + "u03a6")


def test_write_surface_audit_caps_cover_largest_native_module():
    # The native inventory found 265 write-surface matches in one report module.
    assert module.MAX_MATCHES_PER_COMPONENT >= 265
    assert module.MAX_TOTAL_MATCHES >= 1396


def test_value_access_is_not_mislabeled_as_a_write():
    assert module._operation("patientName = ws.Cells(rowIndex, 1).Value") == (
        "cell_value_or_formula_access", None
    )
    assert module._operation("ws.Cells(rowIndex, 1).Value = patientName") == (
        "cell_value_or_formula_write", None
    )


def test_procedure_inventory_returns_names_without_source_text():
    names = module._procedure_names([
        "Option Explicit",
        "Private Sub Worksheet_Change(ByVal Target As Range)",
        "    secretPatientName = Target.Value",
        "End Sub",
        "Public Function ValidateRow() As Boolean",
        "End Function",
    ])
    assert names == ["Worksheet_Change", "ValidateRow"]
    assert "secretPatientName" not in ",".join(names)


def test_unmatched_component_audit_reports_procedure_names_not_source():
    source = inspect.getsource(module.audit)
    assert "NO_PATTERN_MATCH_COMPONENT" in source
    assert "_procedure_names(lines)" in source
    assert "procedure_names" in source


def test_external_file_and_process_side_effects_are_classified():
    assert module._operation('Kill requestPath') == ("external_file_write_or_delete", None)
    assert module._operation('WriteUtf8Text requestPath, payload') == (
        "external_file_write_or_delete", None
    )
    assert module._operation('responseText = ReadUtf8Text(responsePath)') == (
        "external_file_read", None
    )
    assert module._operation('exitCode = RunHidden(commandLine)') == (
        "external_process_launch", None
    )
    assert module._operation('ThisWorkbook.Close SaveChanges:=False') == (
        "workbook_close", None
    )


def test_external_side_effect_patterns_are_scanned():
    lines = [
        "Private Sub Apply()",
        "Kill requestPath",
        "WriteUtf8Text requestPath, payload",
        "exitCode = RunHidden(commandLine)",
        "ThisWorkbook.Close SaveChanges:=False",
        "End Sub",
    ]
    assert [item[2] for item in module._matches(lines)] == lines[1:5]
