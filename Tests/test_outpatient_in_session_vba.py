from rehab_excel.outpatient_in_session_vba import MODULE_CODE, MODULE_NAME


def test_outpatient_writer_saves_directly_in_open_workbook():
    assert MODULE_NAME == "modOutpatientInSession"
    assert "Public Function SaveOutpatientScheduleInWorkbook" in MODULE_CODE
    assert 'ThisWorkbook.Worksheets("PATIENTS")' in MODULE_CODE
    assert 'ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")' in MODULE_CODE
    assert "ThisWorkbook.Save" in MODULE_CODE
    assert "Application.CalculateFull" in MODULE_CODE
    assert "Workbooks.Open" not in MODULE_CODE
    assert "ThisWorkbook.Close" not in MODULE_CODE


def test_outpatient_writer_validates_patient_identity_and_type():
    assert "FindPatientRow" in MODULE_CODE
    assert "IsOutpatientValue" in MODULE_CODE
    assert "Ο ασθενής δεν είναι καταχωρημένος ως εξωτερικός." in MODULE_CODE
