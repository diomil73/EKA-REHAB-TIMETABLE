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


def test_side_effect_scan_does_not_count_procedure_declarations():
    lines = [
        "Private Function RunHidden(ByVal commandLine As String) As Long",
        "    RunHidden = shell.Run(commandLine, 0, True)",
        "End Function",
        "Private Sub WriteUtf8Text(ByVal filePath As String, ByVal text As String)",
        "    stream.SaveToFile filePath, 2",
        "End Sub",
    ]
    assert [item[2] for item in module._matches(lines)] == [
        "RunHidden = shell.Run(commandLine, 0, True)",
        "stream.SaveToFile filePath, 2",
    ]


def test_procedure_profiles_are_source_redacted_and_stable():
    lines = [
        "Private Function NormalizeTimeMS(ByVal rawTime As Variant) As String",
        "    If IsDate(rawTime) Then",
        '        NormalizeTimeMS = Format(CDate(rawTime), "hh:mm")',
        "    Else",
        '        NormalizeTimeMS = ""',
        "    End If",
        "End Function",
    ]
    profiles = module._procedure_profiles(lines)
    assert len(profiles) == 1
    profile = profiles[0]
    assert profile["name"] == "NormalizeTimeMS"
    assert profile["body_lines"] == 5
    assert profile["categories"]["branches"] == 3
    assert profile["categories"]["date_time"] == 1
    assert len(profile["source_sha256"]) == 64
    assert "rawTime" not in str(profile)
    assert "hh:mm" not in str(profile)
    assert "NormalizeTimeMS = " not in str(profile)


def test_procedure_profiles_are_stable_across_identical_source():
    lines = [
        "Private Sub Worksheet_Change(ByVal Target As Range)",
        "    If Target.CountLarge > 1 Then Exit Sub",
        "End Sub",
    ]
    first = module._procedure_profiles(lines)
    second = module._procedure_profiles(list(lines))
    assert first == second
    assert first[0]["categories"]["branches"] == 1


def test_unmatched_component_audit_emits_only_profile_metadata():
    source = inspect.getsource(module.audit)
    assert "PROCEDURE_PROFILE" in source
    assert "component_sha256=" in source
    assert "_procedure_profiles(lines)" in source


def test_possible_vba_call_edges_ignore_literals_and_comments():
    lines = [
        "Private Sub Worksheet_Change(ByVal Target As Range)",
        "    If Target.CountLarge > 1 Then Exit Sub",
        "    Call RefreshTherapistDaily",
        '    message = "Call RestoreReturnedTherapists"',
        "    ' Call ColorReplacementMarks",
        "    If Not Intersect(Target, Range(\"A1\")) Is Nothing Then RefreshMaster",
        "End Sub",
    ]
    assert module._possible_call_edges(
        lines,
        {
            "RefreshTherapistDaily",
            "RestoreReturnedTherapists",
            "ColorReplacementMarks",
            "RefreshMaster",
        },
    ) == [
        ("Worksheet_Change", 3, "RefreshTherapistDaily"),
        ("Worksheet_Change", 6, "RefreshMaster"),
    ]


def test_possible_vba_call_edges_are_candidates_not_execution_proofs():
    lines = [
        "Private Function NormalizeTimeMS(ByVal rawValue As Variant)",
        "    NormalizeTimeMS = rawValue",
        "End Function",
        "Private Sub NextStep()",
        '    s = "NormalizeTimeMS"',
        "End Sub",
    ]
    assert module._possible_call_edges(lines, {"NormalizeTimeMS"}) == []


def test_call_edge_auditor_does_not_print_source_lines():
    source = inspect.getsource(module.audit)
    assert "POSSIBLE_CALL_EDGE" in source
    assert "_possible_call_edges" in source
    assert "_safe_label(caller)" in source
    assert "_safe_label(callee)" in source
    assert "f\"  POSSIBLE_CALL_EDGE caller=" in source


def test_component_sha256_uses_newlines_not_literal_backslash_n():
    source = inspect.getsource(module.audit)
    assert 'sha256("\\n".join(lines).encode("utf-8")).hexdigest()' in source
    assert 'sha256("\\\\n".join(lines).encode("utf-8")).hexdigest()' not in source


def test_normalizer_call_search_covers_modules_with_write_matches():
    source = inspect.getsource(module.audit)
    assert 'if name.casefold() in {"normalizetimems", "normalizepresencems"}' in source
    assert "lines, normalizer_names" in source
    assert "POSSIBLE_CALL_EDGE" in source
