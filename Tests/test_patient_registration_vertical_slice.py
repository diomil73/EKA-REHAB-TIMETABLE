from rehab_excel.patient_registration_vertical_slice import (
    _vertical_form_code,
    _vertical_module_code,
)


def test_vertical_patient_registration_persists_responsible_doctor():
    form_code = _vertical_form_code()
    module_code = _vertical_module_code()

    assert "txtResponsibleDoctor.Text" in form_code
    assert "ByVal responsibleDoctor As String" in module_code
    assert "patients.Cells(targetRow, doctorCol).Value = Trim$(responsibleDoctor)" in module_code
    assert 'If IsOutpatient(patientType) Then\n        patients.Cells(targetRow, doctorCol).Value = ""' in module_code


def test_vertical_patient_registration_keeps_v36_in_session_navigation():
    form_code = _vertical_form_code()

    assert "GoToMaster" in form_code
    assert "ThisWorkbook.Close" not in form_code
    assert "Application.Quit" not in form_code


def test_vertical_master_links_are_blank_safe_instead_of_rendering_zeroes():
    module_code = _vertical_module_code()

    assert '"=IF(PATIENT_PLANNER!F" & plannerRow & "="""","""",PATIENT_PLANNER!F" & plannerRow & ")"' in module_code
    assert '"=IF(PATIENT_PLANNER!I" & plannerRow & "="""","""",PATIENT_PLANNER!I" & plannerRow & ")"' in module_code
    assert '"=IF(PATIENT_PLANNER!N" & plannerRow & "="""","""",PATIENT_PLANNER!N" & plannerRow & ")"' in module_code
    assert 'ws.Cells(targetRow, 4).Formula = "=PATIENT_PLANNER!F" & plannerRow' not in module_code


def test_master_doctor_line_matches_approved_secondary_visual_spec():
    module_code = _vertical_module_code()

    assert "Private Function DoctorSeparator() As String" in module_code
    assert "ChrW$(9472)" in module_code
    assert "Private Sub WritePatientDoctorCell" in module_code
    assert "displayName & vbLf & separator & vbLf & doctor" in module_code
    assert '"✚ " & doctor' not in module_code
    assert "Private Sub ApplyDoctorLineStyle" in module_code
    assert '.Name = "Segoe UI"' in module_code
    assert ".Size = 10" in module_code
    assert ".Size = 9" in module_code
    assert ".Italic = True" in module_code
    assert ".Color = RGB(92, 64, 120)" in module_code
    assert ".Color = RGB(180, 185, 195)" in module_code
    assert "ApplyDoctorLineStyle ws.Cells(targetRow, 3), displayName, doctor" in module_code
    assert "ws.Rows(rowIndex).RowHeight = 72" in module_code
