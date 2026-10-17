from rehab_excel.scheduling_conflict_popup_vba import (
    SCHEDULING_CONFLICT_FORM_CODE,
    SCHEDULING_CONFLICT_FORM_NAME,
    SCHEDULING_CONFLICT_MODULE_CODE,
    SCHEDULING_CONFLICT_MODULE_NAME,
)


def test_conflict_popup_uses_explicit_accept_and_change_buttons():
    assert SCHEDULING_CONFLICT_FORM_NAME == "frmSchedulingConflict"
    assert SCHEDULING_CONFLICT_MODULE_NAME == "modSchedulingConflict"
    assert '.Caption = "Δεκτό"' in SCHEDULING_CONFLICT_FORM_CODE
    assert '.Caption = "Αλλαγή"' in SCHEDULING_CONFLICT_FORM_CODE


def test_change_is_safe_default_and_cancel_path():
    assert ".Default = True" in SCHEDULING_CONFLICT_FORM_CODE
    assert ".Cancel = True" in SCHEDULING_CONFLICT_FORM_CODE
    assert "Accepted = False" in SCHEDULING_CONFLICT_FORM_CODE


def test_popup_names_therapist_time_and_both_patients():
    assert "therapistName" in SCHEDULING_CONFLICT_FORM_CODE
    assert "existingPatientName" in SCHEDULING_CONFLICT_FORM_CODE
    assert "newPatientName" in SCHEDULING_CONFLICT_FORM_CODE
    assert '" στις " & timeText' in SCHEDULING_CONFLICT_FORM_CODE


def test_conflict_guard_rings_and_requires_explicit_override():
    assert "Beep" in SCHEDULING_CONFLICT_MODULE_CODE
    assert "ConfirmTherapistDoubleBooking" in SCHEDULING_CONFLICT_MODULE_CODE
    assert "ConfirmTherapistDoubleBooking = dlg.Accepted" in SCHEDULING_CONFLICT_MODULE_CODE


def test_popup_is_modal():
    assert "dlg.Show vbModal" in SCHEDULING_CONFLICT_MODULE_CODE
