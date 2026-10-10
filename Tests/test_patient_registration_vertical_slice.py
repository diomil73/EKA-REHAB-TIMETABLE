from rehab_excel.patient_registration_vertical_slice import (
    _vertical_form_code,
    _vertical_module_code,
)


def test_vertical_patient_registration_requires_and_persists_responsible_doctor():
    form_code = _vertical_form_code()
    module_code = _vertical_module_code()

    assert "txtResponsibleDoctor.Text" in form_code
    assert "Ο υπεύθυνος γιατρός είναι υποχρεωτικός." in form_code
    assert "ByVal responsibleDoctor As String" in module_code
    assert "patients.Cells(targetRow, doctorCol).Value = Trim$(responsibleDoctor)" in module_code


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
