from rehab_excel.student_registration_form_vba import (
    STUDENT_FORM_CODE,
    STUDENT_FORM_NAME,
)


def test_student_form_name_is_stable():
    assert STUDENT_FORM_NAME == "frmNewStudent"


def test_student_form_collects_required_fields_without_second_number():
    for marker in (
        "txtStudentID",
        "txtDisplayName",
        "txtPlacementStart",
        "txtPlacementEnd",
        "cboSupervisor",
        "chkReplacement",
        "chkRobotic",
    ):
        assert marker in STUDENT_FORM_CODE
    assert "txtStudentNumber" not in STUDENT_FORM_CODE
    assert "student_number" not in STUDENT_FORM_CODE


def test_student_id_is_locked_and_sent_blank_for_backend_allocation():
    assert 'txtStudentID.Text = "Αυτόματο"' in STUDENT_FORM_CODE
    assert "txtStudentID.Locked = True" in STUDENT_FORM_CODE
    assert "txtStudentID.TabStop = False" in STUDENT_FORM_CODE
    assert 'q & "student_id" & q & ":" & q & q' in STUDENT_FORM_CODE
    assert "Το Student ID είναι υποχρεωτικό" not in STUDENT_FORM_CODE


def test_student_form_uses_clear_save_button():
    assert '.Caption = "Αποθήκευση"' in STUDENT_FORM_CODE
    assert "Έλεγχος και preview" not in STUDENT_FORM_CODE


def test_student_form_loads_supervisors_from_settings():
    assert 'ThisWorkbook.Worksheets("SETTINGS")' in STUDENT_FORM_CODE
    assert "cboSupervisor.AddItem" in STUDENT_FORM_CODE


def test_student_form_calls_shared_registration_bridge():
    assert "registration_bridge_cli.py" in STUDENT_FORM_CODE
    assert 'q & "action" & q & ":" & q & "new_student" & q' in STUDENT_FORM_CODE
    assert 'q & "student_id" & q' in STUDENT_FORM_CODE
    assert 'q & "replacement_capable" & q' in STUDENT_FORM_CODE
    assert 'q & "robotic_capable" & q' in STUDENT_FORM_CODE


def test_student_form_uses_utf8_and_preview_source():
    assert "ThisWorkbook.FullName" in STUDENT_FORM_CODE
    assert "ThisWorkbook.Path" in STUDENT_FORM_CODE
    assert 'CreateObject("ADODB.Stream")' in STUDENT_FORM_CODE
    assert 'stream.Charset = "utf-8"' in STUDENT_FORM_CODE


def test_student_form_wires_authoritative_commit_worker():
    assert "StartAuthoritativeCommit" in STUDENT_FORM_CODE
    assert "authoritative_commit_worker_cli.py" in STUDENT_FORM_CODE
    assert 'JsonStringValue(responseText, "source_sha256_before")' in STUDENT_FORM_CODE
    assert 'JsonStringValue(responseText, "output_path")' in STUDENT_FORM_CODE
    assert 'q & "expected_source_sha256" & q' in STUDENT_FORM_CODE
    assert 'q & "preview_path" & q' in STUDENT_FORM_CODE
    assert 'q & "reopen" & q & ":true,"' in STUDENT_FORM_CODE


def test_student_form_starts_worker_async_then_closes_authoritative_workbook():
    save_proc_end = STUDENT_FORM_CODE.index("CleanUp:")
    save_proc = STUDENT_FORM_CODE[:save_proc_end]
    start_pos = save_proc.index("If Not StartAuthoritativeCommit(")
    close_pos = save_proc.index("ThisWorkbook.Close SaveChanges:=True")
    assert start_pos < close_pos
    assert 'CreateObject("WScript.Shell").Run commandLine, 0, False' in STUDENT_FORM_CODE
    assert "Unload Me" in save_proc


def test_student_form_sets_returned_student_id_before_close():
    id_pos = STUDENT_FORM_CODE.index(
        'txtStudentID.Text = JsonStringValue(responseText, "subject_key")'
    )
    start_pos = STUDENT_FORM_CODE.index("If Not StartAuthoritativeCommit(")
    assert id_pos < start_pos
