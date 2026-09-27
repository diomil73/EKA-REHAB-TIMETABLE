from rehab_excel.outpatient_schedule_form_vba import FORM_CODE, FORM_NAME, MACRO_NAME


def test_form_and_macro_names_are_stable():
    assert FORM_NAME == "frmOutpatientSchedule"
    assert MACRO_NAME == "ShowOutpatientScheduleForm"


def test_form_loads_only_outpatients_and_real_settings():
    assert 'ThisWorkbook.Worksheets("PATIENTS")' in FORM_CODE
    assert 'patientType = Trim$(CStr(ws.Cells(rowIndex, 6).Value))' in FORM_CODE
    assert 'patientType = "Εξωτερικός"' in FORM_CODE
    assert 'ThisWorkbook.Worksheets("SETTINGS")' in FORM_CODE
    assert "Cells(ws.Rows.Count, 1)" in FORM_CODE
    assert "Cells(ws.Rows.Count, 2)" in FORM_CODE
    assert "Cells(ws.Rows.Count, 4)" in FORM_CODE
    assert "Cells(ws.Rows.Count, 5)" in FORM_CODE


def test_form_does_not_offer_robotic_treatment():
    assert 'valueText <> "Ρομποτικό"' in FORM_CODE


def test_form_calls_json_bridge_and_uses_preview_only_source():
    assert "outpatient_schedule_bridge_cli.py" in FORM_CODE
    assert "ThisWorkbook.FullName" in FORM_CODE
    assert "ThisWorkbook.Path" in FORM_CODE
    assert 'q & "overwrite" & q & ":true,"' in FORM_CODE


def test_form_sends_patient_treatment_time_days_and_therapist():
    for key in ("patient_id", "treatment", "time", "days", "therapist"):
        assert f'q & "{key}" & q' in FORM_CODE


def test_form_uses_utf8_and_bom_tolerant_cli_contract():
    assert 'CreateObject("ADODB.Stream")' in FORM_CODE
    assert 'stream.Charset = "utf-8"' in FORM_CODE


def test_form_has_final_seven_suggestion_modes_in_order():
    labels = (
        "Αλλαγή ώρας",
        "Αλλαγή ημερών",
        "Αλλαγή ώρας και ημερών",
        "Αλλαγή θεραπευτή",
        "Αλλαγή θεραπευτή και ώρας",
        "Αλλαγή θεραπευτή και ημερών",
        "Αλλαγή θεραπευτή, ώρας και ημερών",
    )
    positions = [FORM_CODE.index(f'cboSuggestMode.AddItem "{label}"') for label in labels]
    assert positions == sorted(positions)


def test_form_calls_suggestion_bridge_and_can_apply_selected_result():
    assert "outpatient_schedule_suggestion_bridge_cli.py" in FORM_CODE
    assert "BuildSuggestionRequestJson" in FORM_CODE
    assert 'q & "mode" & q' in FORM_CODE
    assert 'q & "limit" & q & ":5"' in FORM_CODE
    assert 'JsonStringValue(responseText, "suggestion_lines")' in FORM_CODE
    assert 'parts = Split(CStr(lstSuggestions.Value), " | ")' in FORM_CODE
    assert "cboTherapist.Value = Trim$(parts(0))" in FORM_CODE
    assert "cboTime.Value = Trim$(parts(1))" in FORM_CODE
    assert "cboDays.Value = Trim$(parts(2))" in FORM_CODE
