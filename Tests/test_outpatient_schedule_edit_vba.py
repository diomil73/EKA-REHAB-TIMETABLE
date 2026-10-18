from rehab_excel.outpatient_schedule_edit_vba import (
    EDIT_MACRO_NAME,
    EDIT_MODULE_CODE,
    patch_outpatient_schedule_form_code,
)
from rehab_excel.outpatient_schedule_form_vba import FORM_CODE


def test_edit_macro_preloads_existing_outpatient_schedule_row():
    assert EDIT_MACRO_NAME == "EditOutpatientSchedule"
    assert 'ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")' in EDIT_MODULE_CODE
    assert 'OutpatientEditTargetBaseEntryId = "outpatient:" & CStr(rowIndex)' in EDIT_MODULE_CODE
    assert '.cboPatient.Value = patientId & " | " & patientName' in EDIT_MODULE_CODE
    assert ".cboTreatment.Value = treatment" in EDIT_MODULE_CODE
    assert ".cboTime.Value = timeText" in EDIT_MODULE_CODE
    assert ".cboDays.Value = dayPattern" in EDIT_MODULE_CODE
    assert ".cboTherapist.Value = therapistId" in EDIT_MODULE_CODE


def test_edit_patch_sends_target_base_entry_id_through_existing_bridge():
    patched = patch_outpatient_schedule_form_code(FORM_CODE)

    assert 'OutpatientEditTargetBaseEntryId = ""' in patched
    assert 'q & "target_base_entry_id" & q' in patched
    assert "JsonEscape(OutpatientEditTargetBaseEntryId)" in patched
    assert "outpatient_schedule_bridge_cli.py" in patched
    assert "StartAuthoritativeCommit" in patched


def test_edit_patch_accepts_vbide_crlf_code():
    patched = patch_outpatient_schedule_form_code(FORM_CODE.replace("\n", "\r\n"))

    assert 'OutpatientEditTargetBaseEntryId = ""' in patched
    assert patched.count('q & "target_base_entry_id" & q') >= 2
    assert "\r" not in patched


def test_edit_patch_is_idempotent():
    once = patch_outpatient_schedule_form_code(FORM_CODE)
    twice = patch_outpatient_schedule_form_code(once)
    assert twice == once
