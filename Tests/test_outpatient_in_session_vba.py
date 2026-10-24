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


def test_outpatient_writer_refreshes_current_therapist_daily_cell_in_session():
    assert "RefreshTherapistDailyForCurrentDate patientName, therapistName, timeText, dayPattern, treatment" in MODULE_CODE
    assert 'ThisWorkbook.Worksheets("DAILY_INPUT")' in MODULE_CODE
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY")' in MODULE_CODE
    assert 'dailyInput.Range("B2").Value' in MODULE_CODE
    assert "DayPatternIncludesDate" in MODULE_CODE
    assert "FindTherapistDailyCell" in MODULE_CODE
    assert "targetCell.Value = previousText & vbLf & patientName" in MODULE_CODE


def test_outpatient_daily_refresh_preserves_existing_cell_when_override_is_accepted():
    assert "If Len(Trim$(previousText)) = 0 Then" in MODULE_CODE
    assert "targetCell.Interior.Color = RGB(221, 235, 247)" in MODULE_CODE
    assert "targetCell.Value = previousText & vbLf & patientName" in MODULE_CODE


def test_therapist_daily_mapping_normalizes_accents_spaces_and_time_text():
    assert "Private Function NormalizeProviderName" in MODULE_CODE
    assert 'Replace(text, "ά", "α")' in MODULE_CODE
    assert 'Replace(text, "έ", "ε")' in MODULE_CODE
    assert 'Replace(text, " ", "")' in MODULE_CODE
    assert "Private Function TimeKey" in MODULE_CODE
    assert 'Replace(Trim$(CStr(rawValue)), ".", ":")' in MODULE_CODE
    assert "NormalizeProviderName(valueText) = targetProvider" in MODULE_CODE


def test_in_session_daily_refresh_marks_robotic_patient_text_orange():
    assert "Private Function IsRoboticTreatment" in MODULE_CODE
    assert "RGB(255, 140, 0)" in MODULE_CODE
    assert "targetCell.Characters(lineStart, Len(patientName)).Font.Bold = True" in MODULE_CODE
