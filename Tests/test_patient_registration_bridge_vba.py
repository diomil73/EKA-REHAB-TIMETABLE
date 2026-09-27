from rehab_excel.patient_registration_bridge_vba import (
    FORM_BRIDGE_CODE,
    PATIENT_FORM_NAME,
)


def test_bridge_targets_existing_patient_form():
    assert PATIENT_FORM_NAME == "frmNewPatient"
    assert "Private Sub cmdSave_Click()" in FORM_BRIDGE_CODE


def test_bridge_calls_json_cli_and_keeps_preview_only_source_path():
    assert "registration_bridge_cli.py" in FORM_BRIDGE_CODE
    assert 'q & "source_path" & q' in FORM_BRIDGE_CODE
    assert "ThisWorkbook.FullName" in FORM_BRIDGE_CODE
    assert 'q & "preview_dir" & q' in FORM_BRIDGE_CODE
    assert "ThisWorkbook.Path" in FORM_BRIDGE_CODE
    assert 'q & "overwrite" & q & ":true,"' in FORM_BRIDGE_CODE


def test_bridge_sends_patient_type_mrn_and_auto_id_request():
    assert 'q & "patient_id" & q & ":" & q & q' in FORM_BRIDGE_CODE
    assert 'q & "patient_type" & q' in FORM_BRIDGE_CODE
    assert 'q & "hospital_mrn" & q' in FORM_BRIDGE_CODE
    assert "txtHospitalMRN.Text" in FORM_BRIDGE_CODE
    assert "txtDisplayName.Text" in FORM_BRIDGE_CODE


def test_bridge_uses_vba_safe_quotes_not_c_style_json_literals():
    assert 'Chr$(34) & "ok" & Chr$(34) & ": false"' in FORM_BRIDGE_CODE
    assert r'\"ok\": false' not in FORM_BRIDGE_CODE
    assert 'q = Chr$(34)' in FORM_BRIDGE_CODE


def test_bridge_forces_inpatient_only_fields_blank_for_outpatient():
    assert 'If patientType = "Εξωτερικός" Then' in FORM_BRIDGE_CODE
    assert 'roomValue = ""' in FORM_BRIDGE_CODE
    assert 'statusValue = ""' in FORM_BRIDGE_CODE
    assert 'infectiousValue = "false"' in FORM_BRIDGE_CODE


def test_bridge_uses_utf8_for_greek_json_roundtrip():
    assert 'CreateObject("ADODB.Stream")' in FORM_BRIDGE_CODE
    assert 'stream.Charset = "utf-8"' in FORM_BRIDGE_CODE


def test_bridge_surfaces_backend_error_and_returned_patient_id():
    assert 'JsonStringValue(responseText, "error")' in FORM_BRIDGE_CODE
    assert 'JsonStringValue(responseText, "subject_key")' in FORM_BRIDGE_CODE
    assert 'JsonStringValue(responseText, "output_path")' in FORM_BRIDGE_CODE
