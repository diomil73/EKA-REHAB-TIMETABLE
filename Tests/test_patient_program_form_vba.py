from rehab_excel.patient_program_form_vba import (
    PATIENT_PROGRAM_FORM_CODE,
    PATIENT_PROGRAM_MODULE_CODE,
)


def test_program_form_loads_configuration_from_patient_program_bridge():
    runtime = PATIENT_PROGRAM_MODULE_CODE
    form = PATIENT_PROGRAM_FORM_CODE
    assert "patient_program_bridge.py" in runtime
    assert '"config-vba "' in runtime
    assert "LoadPatientProgramConfig" in runtime
    assert 'Case "TREATMENT"' in form
    assert 'Case "PROVIDER"' in form
    assert 'Case "DAY_PATTERN"' in form
    assert 'Case "STANDARD_SLOT"' in form
    assert 'Case "RECLINED_SLOT"' in form


def test_program_form_uses_checkbox_style_multiselect_treatment_list():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "lstTreatments.MultiSelect = 1" in code
    assert "lstTreatments.ListStyle = 1" in code
    assert "lstTreatments.AddItem treatment" in code
    assert "cboTreatment" not in code


def test_program_form_keeps_independent_draft_fields_per_treatment():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "mDraftDays" in code
    assert "mDraftTimes" in code
    assert "mDraftProviders" in code
    assert "SaveCurrentDraft" in code
    assert "LoadCurrentTreatmentEditor" in code
    assert "Τα συμπληρωμένα πεδία θεωρούνται χειροκίνητες επιλογές" in code


def test_program_form_exposes_recliner_adjacency_preference():
    code = PATIENT_PROGRAM_FORM_CODE
    assert 'StrComp(treatment, "Ανακλινόμενο", vbTextCompare) = 0' in code
    assert "mReclinedSlots" in code
    assert "το Ανακλινόμενο θα τοποθετείται πριν ή μετά τη ΦΘ" in code


def test_program_form_does_not_write_workbook_before_final_commit_exists():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "ThisWorkbook.Save" not in code
    assert "PATIENT_PLANNER" not in code
    assert "SavePatientRecurringAssignment" not in code
    assert "cmdSaveProgram.Enabled = False" in code


def test_program_form_stays_modeless_and_guarded_against_double_open():
    code = PATIENT_PROGRAM_MODULE_CODE
    start = code.index("Public Sub OpenPatientScheduling")
    procedure = code[start:]
    assert "If frmPatientScheduling.Visible Then Exit Sub" in procedure
    assert "frmPatientScheduling.Show vbModeless" in procedure
