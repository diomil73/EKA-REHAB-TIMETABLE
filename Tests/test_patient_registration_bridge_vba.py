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


def test_bridge_keeps_outpatient_status_but_blanks_inpatient_only_fields():
    assert 'statusValue = Trim$(cboStatus.Value)' in FORM_BRIDGE_CODE
    assert 'If patientType = "Εξωτερικός" Then' in FORM_BRIDGE_CODE
    assert 'roomValue = ""' in FORM_BRIDGE_CODE
    assert 'infectiousValue = "false"' in FORM_BRIDGE_CODE


def test_bridge_uses_utf8_for_greek_json_roundtrip():
    assert 'CreateObject("ADODB.Stream")' in FORM_BRIDGE_CODE
    assert 'stream.Charset = "utf-8"' in FORM_BRIDGE_CODE


def test_bridge_surfaces_backend_error_and_returned_patient_id():
    assert 'JsonStringValue(responseText, "error")' in FORM_BRIDGE_CODE
    assert 'JsonStringValue(responseText, "subject_key")' in FORM_BRIDGE_CODE
    assert 'JsonStringValue(responseText, "output_path")' in FORM_BRIDGE_CODE


def test_bridge_wires_authoritative_commit_worker():
    assert "StartAuthoritativeCommit" in FORM_BRIDGE_CODE
    assert "authoritative_commit_worker_cli.py" in FORM_BRIDGE_CODE
    assert 'JsonStringValue(responseText, "source_sha256_before")' in FORM_BRIDGE_CODE
    assert 'JsonStringValue(responseText, "output_path")' in FORM_BRIDGE_CODE
    assert 'q & "expected_source_sha256" & q' in FORM_BRIDGE_CODE
    assert 'q & "preview_path" & q' in FORM_BRIDGE_CODE
    assert 'q & "reopen" & q & ":true,"' in FORM_BRIDGE_CODE


def test_bridge_starts_worker_async_then_closes_authoritative_workbook():
    save_proc_end = FORM_BRIDGE_CODE.index("CleanUp:")
    save_proc = FORM_BRIDGE_CODE[:save_proc_end]
    start_pos = save_proc.index("If Not StartAuthoritativeCommit(")
    close_pos = save_proc.index("ThisWorkbook.Close SaveChanges:=True")
    assert start_pos < close_pos
    assert 'CreateObject("WScript.Shell").Run commandLine, 0, False' in FORM_BRIDGE_CODE
    assert "Unload Me" in save_proc
