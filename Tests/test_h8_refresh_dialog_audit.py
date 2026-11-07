from tools.audit_h8_refresh_dialogs import REFRESH, dialog_audit, extract_procedures


def test_extract_vba_procedures_and_report_subordinate_dialogs():
    bodies = {
        "Module3": """Public Sub RefreshAll()
    Call RefreshTherapistDaily
    Call RefreshTherapistAttendanceAndConflicts
    MsgBox "Όλα ενημερώθηκαν."
End Sub
""",
        "modOps": """Public Sub RefreshTherapistDaily()
    Call InternalUpdate
End Sub
Private Sub InternalUpdate()
    MsgBox "Ενδιάμεσο"
End Sub
""",
    }
    for index, procedure in enumerate(REFRESH[1:], start=1):
        bodies[f"safe{index}"] = f"Public Sub {procedure}()\nEnd Sub\n"
    procs = extract_procedures(bodies)
    assert "refreshall" in procs
    assert "internalupdate" in procs
    records, missing = dialog_audit(bodies)
    assert not missing
    assert any("Ενδιάμεσο" in item for item in records)
    assert any("Όλα ενημερώθηκαν" in item for item in records)


def test_missing_refresh_dependency_is_fail_closed():
    records, missing = dialog_audit({"Module3": "Public Sub RefreshAll()\nEnd Sub"})
    assert records
    assert set(missing) == set(REFRESH)


def test_duplicate_procedure_is_safety_error():
    import pytest

    with pytest.raises(ValueError, match="Ambiguous"):
        dialog_audit({
            "A": "Sub RefreshAll()\nEnd Sub",
            "B": "Public Sub RefreshAll()\nEnd Sub",
        })
