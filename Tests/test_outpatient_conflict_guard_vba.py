from rehab_excel.outpatient_conflict_guard_vba import patch_outpatient_schedule_form_code
from rehab_excel.outpatient_schedule_form_vba import FORM_CODE


def test_conflict_guard_patches_save_flow_for_explicit_override():
    patched = patch_outpatient_schedule_form_code(FORM_CODE)

    assert "BuildRequestJson(False)" in patched
    assert "BuildRequestJson(True)" in patched
    assert 'JsonStringValue(responseText, "conflict_type") = "therapist_double_booking"' in patched
    assert "ConfirmTherapistDoubleBooking" in patched
    assert 'JsonStringValue(responseText, "existing_patient_name")' in patched
    assert 'JsonStringValue(responseText, "new_patient_name")' in patched
    assert 'q & "allow_double_booking" & q' in patched


def test_conflict_guard_keeps_change_as_safe_exit_path():
    patched = patch_outpatient_schedule_form_code(FORM_CODE)

    assert "If Not accepted Then GoTo CleanUp" in patched
    assert patched.count("BuildRequestJson(True)") == 1


def test_outpatient_save_is_in_session_and_never_closes_workbook():
    patched = patch_outpatient_schedule_form_code(FORM_CODE)

    assert "SaveOutpatientScheduleInWorkbook" in patched
    assert "ThisWorkbook.Close" not in patched
    assert "StartAuthoritativeCommit(" not in patched
    assert "BuildAuthoritativeCommitJson(" not in patched
    assert "ResolveAuthoritativeCommitWorkerScript" not in patched
    assert "χωρίς κλείσιμο ή επανεκκίνηση" in patched


def test_outpatient_in_session_patch_is_idempotent():
    patched_once = patch_outpatient_schedule_form_code(FORM_CODE)
    patched_twice = patch_outpatient_schedule_form_code(patched_once)

    assert patched_twice == patched_once
