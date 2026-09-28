from rehab_excel.student_registration_form_vba import (
    STUDENT_FORM_CODE,
    STUDENT_FORM_NAME,
)


def test_student_form_name_is_stable():
    assert STUDENT_FORM_NAME == "frmNewStudent"


def test_student_form_collects_required_fields():
    for marker in (
        "txtStudentID",
        "txtDisplayName",
        "txtStudentNumber",
        "txtPlacementStart",
        "txtPlacementEnd",
        "cboSupervisor",
        "chkReplacement",
        "chkRobotic",
    ):
        assert marker in STUDENT_FORM_CODE


def test_student_form_loads_supervisors_from_settings():
    assert 'ThisWorkbook.Worksheets("SETTINGS")' in STUDENT_FORM_CODE
    assert "cboSupervisor.AddItem" in STUDENT_FORM_CODE


def test_student_form_calls_shared_registration_bridge():
    assert "registration_bridge_cli.py" in STUDENT_FORM_CODE
    assert 'q & "action" & q & ":" & q & "new_student" & q' in STUDENT_FORM_CODE
    assert 'q & "student_id" & q' in STUDENT_FORM_CODE
    assert 'q & "student_number" & q' in STUDENT_FORM_CODE
    assert 'q & "replacement_capable" & q' in STUDENT_FORM_CODE
    assert 'q & "robotic_capable" & q' in STUDENT_FORM_CODE


def test_student_form_uses_utf8_and_preview_source():
    assert "ThisWorkbook.FullName" in STUDENT_FORM_CODE
    assert "ThisWorkbook.Path" in STUDENT_FORM_CODE
    assert 'CreateObject("ADODB.Stream")' in STUDENT_FORM_CODE
    assert 'stream.Charset = "utf-8"' in STUDENT_FORM_CODE


def test_student_form_closes_after_success():
    success_marker = 'MsgBox "Δημιουργήθηκε ασφαλές preview εγγραφής."'
    success_pos = STUDENT_FORM_CODE.index(success_marker)
    unload_pos = STUDENT_FORM_CODE.index("Unload Me", success_pos)
    cleanup_pos = STUDENT_FORM_CODE.index("CleanUp:", success_pos)
    assert success_pos < unload_pos < cleanup_pos
