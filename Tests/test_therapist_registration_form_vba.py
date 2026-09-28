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
    assert "Δημιουργήθηκε ασφαλές preview εγγραφής." in THERAPIST_FORM_CODE
