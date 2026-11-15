"""Synthetic-only tests of read-only afternoon VBA form patch source."""
from pathlib import Path

import pytest

from rehab_excel.afternoon_patient_readonly_vba import (
    PANEL_LABEL, FORM_HELPER, patch_afternoon_panel_text,
)
from rehab_excel.patient_program_form_vba import (
    PATIENT_PROGRAM_FORM_CODE, PATIENT_PROGRAM_MODULE_CODE,
)


def test_panel_installs_in_existing_form_without_replacing_morning_programme():
    runtime, form = patch_afternoon_panel_text(
        PATIENT_PROGRAM_MODULE_CODE, PATIENT_PROGRAM_FORM_CODE,
    )
    assert "Public Function LoadPatientAfternoonState" in runtime
    assert "afternoon-state-vba " in runtime
    assert "RefreshAfternoonReadOnlyPanel" in form
    assert "SavePatientProgramAtomic" in form
    assert "PATIENT_PLANNER" in form
    assert "LoadProgramConfiguration" in form
    assert "lstTreatments" in form
    assert 'lblAfternoonState.Caption' in form
    assert "Πρωινή ΦΘ" not in form
    assert "    Me.Height = 665" in form
    assert "    Me.Height = 585" not in form


def test_readonly_panel_has_no_afternoon_save_or_worksheet_mutation():
    assert "Worksheets" not in FORM_HELPER
    assert "ThisWorkbook.Save" not in FORM_HELPER
    assert "cmdSave" not in FORM_HELPER
    assert "PM_ROW" in FORM_HELPER
    assert "PM_STATUS" in FORM_HELPER


def test_form_patch_is_idempotent():
    first = patch_afternoon_panel_text(
        PATIENT_PROGRAM_MODULE_CODE, PATIENT_PROGRAM_FORM_CODE
    )
    assert patch_afternoon_panel_text(*first) == first
    assert first[0].count("Public Function LoadPatientAfternoonState") == 1
    assert first[1].count("Private Sub RefreshAfternoonReadOnlyPanel") == 1


def test_unknown_form_revision_fails_closed():
    with pytest.raises(RuntimeError):
        patch_afternoon_panel_text(PATIENT_PROGRAM_MODULE_CODE, "Option Explicit")
    with pytest.raises(RuntimeError):
        patch_afternoon_panel_text("Option Explicit", PATIENT_PROGRAM_FORM_CODE)


def test_existing_conflict_patches_have_same_required_markers():
    from rehab_excel.patient_program_master_render_vba import (
        _patch_program_conflict_validation,
    )
    runtime, form = patch_afternoon_panel_text(
        PATIENT_PROGRAM_MODULE_CODE, PATIENT_PROGRAM_FORM_CODE
    )
    assert "Private Sub ApplyAutomaticCompletion(ByVal response As String)" in form
    assert "Private Sub SavePatientProgramAtomic()" in form
    assert "cmdSaveProgram_Click" in form


def test_installer_runs_only_after_existing_form_and_master_hooks():
    source = (
        Path(__file__).resolve().parents[1]
        / "Python" / "rehab_excel" / "patient_registration_scheduling_slice.py"
    ).read_text(encoding="utf-8")
    assert "install_afternoon_readonly_panel" in source
    assert source.index('stage = "installing MASTER patient-program renderer"') < source.index(
        'stage = "installing read-only afternoon patient panel"'
    ) < source.index('stage = "verifying patient-program VBA replacement"')


def test_ui_installer_creates_separate_label_and_does_not_modify_master():
    source = (
        Path(__file__).resolve().parents[1]
        / "Python" / "rehab_excel" / "afternoon_patient_readonly_vba.py"
    ).read_text(encoding="utf-8")
    assert PANEL_LABEL == "lblAfternoonState"
    assert "Forms.Label.1" in source
    assert "MASTER_SCHEDULE" not in source
    assert 'designer.Controls.Add("Forms.Label.1"' in source


def test_native_installer_verifies_afternoon_vba_and_designer_control():
    source = (
        Path(__file__).resolve().parents[1]
        / "Python" / "rehab_excel" / "patient_registration_scheduling_slice.py"
    ).read_text(encoding="utf-8")
    assert '"RefreshAfternoonReadOnlyPanel"' in source
    assert '"LoadPatientAfternoonState"' in source
    assert '"afternoon-state-vba"' in source
    assert 'Designer.Controls(' in source
    assert '"lblAfternoonState"' in source
    assert '"Me.Height = 665"' in source
    assert 'missing_afternoon_control' in source


@pytest.mark.parametrize("ending", ["\r\n", "\r"])
def test_native_vba_newlines_patch_same_as_repository_source(ending):
    """Excel CodeModule.Lines may use CRLF; both source versions must agree."""
    original = patch_afternoon_panel_text(
        PATIENT_PROGRAM_MODULE_CODE, PATIENT_PROGRAM_FORM_CODE
    )
    runtime_native = PATIENT_PROGRAM_MODULE_CODE.replace("\n", ending)
    form_native = PATIENT_PROGRAM_FORM_CODE.replace("\n", ending)
    assert patch_afternoon_panel_text(runtime_native, form_native) == original


def test_native_newline_patched_source_is_idempotent():
    initial = patch_afternoon_panel_text(
        PATIENT_PROGRAM_MODULE_CODE.replace("\n", "\r\n"),
        PATIENT_PROGRAM_FORM_CODE.replace("\n", "\r\n"),
    )
    assert patch_afternoon_panel_text(
        initial[0].replace("\n", "\r\n"),
        initial[1].replace("\n", "\r\n"),
    ) == initial


def test_malformed_native_form_with_crlf_still_fails_closed():
    with pytest.raises(RuntimeError):
        patch_afternoon_panel_text(
            PATIENT_PROGRAM_MODULE_CODE.replace("\n", "\r\n"),
            "Option Explicit\r\nNoMatchingFormMarker\r\n",
        )



def test_overflow_label_describes_extra_sessions_not_days():
    assert "... και άλλες συνεδρίες." in FORM_HELPER
    assert "... και άλλες ημέρες." not in FORM_HELPER
