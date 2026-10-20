from rehab_excel.outpatient_daily_edit_vba import patch_outpatient_in_session_module_code
from rehab_excel.outpatient_in_session_vba import MODULE_CODE


def test_edit_cleanup_handles_decorated_daily_patient_lines_before_rendering_once():
    patched = patch_outpatient_in_session_module_code(MODULE_CODE)

    assert "RemoveOutpatientRenderFromAffectedDailyCells" in patched
    assert "RemovePatientRenderFromDailyCell" in patched
    assert "IsOutpatientPatientRenderLine" in patched
    assert 'Left$(suffix, 2) = " ["' in patched
    assert 'Left$(suffix, 2) = " ("' in patched
    assert 'Left$(suffix, 3) = " → "' in patched
    assert "RemoveFromTherapistDailyForCurrentDate patientName" not in patched


def test_edit_cleanup_cleans_both_old_and_destination_cells_when_slot_changes():
    patched = patch_outpatient_in_session_module_code(MODULE_CODE)

    assert "RemovePatientRenderFromDailyCell dailySheet, patientName, oldTherapistName, oldTimeText" in patched
    assert "RemovePatientRenderFromDailyCell dailySheet, patientName, newTherapistName, newTimeText" in patched
    assert "NormalizeProviderName(oldTherapistName)" in patched
    assert "TimeKey(oldTimeText)" in patched


def test_daily_edit_patch_is_idempotent():
    once = patch_outpatient_in_session_module_code(MODULE_CODE)
    twice = patch_outpatient_in_session_module_code(once)

    assert twice == once
