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


def test_edit_refreshes_visible_therapist_workload_after_daily_projection_change():
    patched = patch_outpatient_in_session_module_code(MODULE_CODE)

    refresh_pos = patched.index(
        "RefreshTherapistDailyForCurrentDate patientName, therapistName, timeText, dayPattern, treatment"
    )
    workload_pos = patched.index("RenderTherapistWorkloadStrip", refresh_pos)
    save_pos = patched.index("ThisWorkbook.Save", workload_pos)

    assert refresh_pos < workload_pos < save_pos
    assert "On Error Resume Next\n    RenderTherapistWorkloadStrip\n    On Error GoTo SaveError" in patched


def test_daily_edit_patch_is_idempotent():
    once = patch_outpatient_in_session_module_code(MODULE_CODE)
    twice = patch_outpatient_in_session_module_code(once)

    assert twice == once
