from rehab_excel.patient_registration_bridge_vba import (
    BRIDGE_MODULE_NAME,
    FORM_BRIDGE_CODE,
    PATIENT_FORM_NAME,
)


def test_bridge_targets_existing_patient_form():
    assert PATIENT_FORM_NAME == "frmNewPatient"
    assert "Private Sub cmdSave_Click()" in FORM_BRIDGE_CODE


def test_bridge_uses_async_registration_transaction_worker():
    assert "registration_transaction_worker_cli.py" in FORM_BRIDGE_CODE
    assert "registration_bridge_cli.py" not in FORM_BRIDGE_CODE
    assert "authoritative_commit_worker_cli.py" not in FORM_BRIDGE_CODE
    assert 'CreateObject("WScript.Shell").Run commandLine, 0, False' in FORM_BRIDGE_CODE


def test_bridge_sends_source_preview_dir_and_patient_values():
    assert 'q & "source_path" & q' in FORM_BRIDGE_CODE
    assert "ThisWorkbook.FullName" in FORM_BRIDGE_CODE
    assert 'q & "preview_dir" & q' in FORM_BRIDGE_CODE
    assert "ThisWorkbook.Path" in FORM_BRIDGE_CODE
    assert 'q & "overwrite" & q & ":true,"' in FORM_BRIDGE_CODE
    assert 'q & "patient_id" & q & ":" & q & q' in FORM_BRIDGE_CODE
    assert 'q & "patient_type" & q' in FORM_BRIDGE_CODE
    assert 'q & "hospital_mrn" & q' in FORM_BRIDGE_CODE
    assert "txtHospitalMRN.Text" in FORM_BRIDGE_CODE
    assert "txtDisplayName.Text" in FORM_BRIDGE_CODE


def test_bridge_keeps_outpatient_status_but_blanks_inpatient_only_fields():
    assert 'statusValue = Trim$(cboStatus.Value)' in FORM_BRIDGE_CODE
    assert 'If patientType = "Εξωτερικός" Then' in FORM_BRIDGE_CODE
    assert 'roomValue = ""' in FORM_BRIDGE_CODE
    assert 'infectiousValue = "false"' in FORM_BRIDGE_CODE


def test_bridge_uses_utf8_for_greek_json_roundtrip():
    assert 'CreateObject("ADODB.Stream")' in FORM_BRIDGE_CODE
    assert 'stream.Charset = "utf-8"' in FORM_BRIDGE_CODE


def test_bridge_returns_from_form_without_waiting_for_preview_or_commit():
    save_proc_end = FORM_BRIDGE_CODE.index("BridgeError:")
    save_proc = FORM_BRIDGE_CODE[:save_proc_end]
    save_pos = save_proc.index("ThisWorkbook.Save")
    request_pos = save_proc.index("WriteUtf8Text requestPath")
    message_pos = save_proc.index('MsgBox "Η καταχώρηση ξεκίνησε.')
    menu_unload_pos = save_proc.index("Unload frmRegistrationMenu")
    launch_pos = save_proc.index('CreateObject("WScript.Shell").Run commandLine, 0, False')
    unload_pos = save_proc.index("Unload Me")

    assert save_pos < request_pos < message_pos < menu_unload_pos < launch_pos < unload_pos
    assert "Run(commandLine, 0, True)" not in save_proc
    assert "ThisWorkbook.Close" not in save_proc
    assert "Application.OnTime" not in FORM_BRIDGE_CODE


def test_bridge_releases_modal_registration_menu_before_worker_launch():
    save_proc_end = FORM_BRIDGE_CODE.index("BridgeError:")
    save_proc = FORM_BRIDGE_CODE[:save_proc_end]
    assert "Unload frmRegistrationMenu" in save_proc
    assert save_proc.index("Unload frmRegistrationMenu") < save_proc.index(
        'CreateObject("WScript.Shell").Run commandLine, 0, False'
    )


def test_bridge_keeps_standard_module_name_for_compatibility():
    assert BRIDGE_MODULE_NAME == "modPatientRegistrationBridge"
