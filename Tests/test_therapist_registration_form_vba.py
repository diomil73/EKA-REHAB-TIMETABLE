from rehab_excel.therapist_registration_form_vba import (
    THERAPIST_FORM_CODE,
    THERAPIST_FORM_NAME,
)


def test_therapist_form_name_is_stable():
    assert THERAPIST_FORM_NAME == "frmNewTherapist"


def test_form_collects_only_therapist_name():
    assert "txtDisplayName" in THERAPIST_FORM_CODE
    assert "robotic" not in THERAPIST_FORM_CODE.casefold()


def test_form_calls_shared_registration_bridge():
    assert "registration_bridge_cli.py" in THERAPIST_FORM_CODE
    assert 'q & "action" & q & ":" & q & "new_therapist" & q' in THERAPIST_FORM_CODE
    assert 'q & "display_name" & q' in THERAPIST_FORM_CODE
    assert 'q & "overwrite" & q & ":true,"' in THERAPIST_FORM_CODE


def test_form_uses_preview_source_and_utf8_transport():
    assert "ThisWorkbook.FullName" in THERAPIST_FORM_CODE
    assert "ThisWorkbook.Path" in THERAPIST_FORM_CODE
    assert 'CreateObject("ADODB.Stream")' in THERAPIST_FORM_CODE
    assert 'stream.Charset = "utf-8"' in THERAPIST_FORM_CODE


def test_form_reports_backend_success_and_error():
    assert 'JsonStringValue(responseText, "error")' in THERAPIST_FORM_CODE
    assert 'JsonStringValue(responseText, "output_path")' in THERAPIST_FORM_CODE
    assert "Η εγγραφή επαληθεύτηκε και είναι έτοιμη για αποθήκευση." in THERAPIST_FORM_CODE


def test_form_wires_authoritative_commit_worker():
    assert "StartAuthoritativeCommit" in THERAPIST_FORM_CODE
    assert "authoritative_commit_worker_cli.py" in THERAPIST_FORM_CODE
    assert 'JsonStringValue(responseText, "source_sha256_before")' in THERAPIST_FORM_CODE
    assert 'JsonStringValue(responseText, "output_path")' in THERAPIST_FORM_CODE
    assert 'q & "expected_source_sha256" & q' in THERAPIST_FORM_CODE
    assert 'q & "preview_path" & q' in THERAPIST_FORM_CODE
    assert 'q & "reopen" & q & ":true,"' in THERAPIST_FORM_CODE


def test_form_starts_worker_async_then_closes_authoritative_workbook():
    save_proc_end = THERAPIST_FORM_CODE.index("CleanUp:")
    save_proc = THERAPIST_FORM_CODE[:save_proc_end]
    start_pos = save_proc.index("If Not StartAuthoritativeCommit(")
    close_pos = save_proc.index("ThisWorkbook.Close SaveChanges:=True")
    assert start_pos < close_pos
    assert 'CreateObject("WScript.Shell").Run commandLine, 0, False' in THERAPIST_FORM_CODE
    assert "Unload Me" in save_proc
