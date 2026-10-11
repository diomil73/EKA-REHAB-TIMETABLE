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


def test_program_form_uses_one_active_treatment_editor_at_a_time():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "lstTreatments.MultiSelect = 0" in code
    assert "lstTreatments.ListStyle = 1" in code
    assert "lstTreatments.AddItem treatment" in code
    assert "mCurrentTreatment = newTreatment" in code
    assert "mIncluded(newTreatment) = True" in code
    assert "cboTreatment" not in code


def test_program_form_shows_each_treatments_draft_values_in_summary_columns():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "lstTreatments.ColumnCount = 5" in code
    assert "lstTreatments.ColumnWidths" in code
    assert "Private Sub RefreshTreatmentRow" in code
    assert 'lstTreatments.List(i, 1) = "✓"' in code
    assert "lstTreatments.List(i, 2) = DraftValue(mDraftDays, treatment)" in code
    assert "lstTreatments.List(i, 3) = DraftValue(mDraftTimes, treatment)" in code
    assert "lstTreatments.List(i, 4) = DraftValue(mDraftProviders, treatment)" in code
    assert "RefreshTreatmentRow mCurrentTreatment" in code
    assert "RefreshAllTreatmentRows" in code


def test_program_form_keeps_included_treatments_separate_from_active_selection():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "Private mIncluded As Object" in code
    assert 'Set mIncluded = CreateObject("Scripting.Dictionary")' in code
    assert "Private Function IsIncluded" in code
    assert "If IsIncluded(treatment) Then" in code
    assert "If lstTreatments.Selected(i) Then" not in code


def test_program_form_keeps_independent_draft_fields_per_treatment():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "mDraftDays" in code
    assert "mDraftTimes" in code
    assert "mDraftProviders" in code
    assert "SaveCurrentDraft" in code
    assert "LoadCurrentTreatmentEditor" in code
    assert "mIncluded(mCurrentTreatment) = True" in code


def test_program_form_makes_current_treatment_explicit_in_editor_labels():
    code = PATIENT_PROGRAM_FORM_CODE
    assert 'lblEditing.Caption = "Ρυθμίσεις θεραπείας: " & treatment' in code
    assert 'lblDays.Caption = "Ημέρες για " & treatment' in code
    assert 'lblTime.Caption = "Ώρα για " & treatment' in code
    assert 'lblProvider.Caption = "Θεραπευτής / provider για " & treatment' in code


def test_program_form_populates_time_and_provider_choices_for_active_treatment():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "For Each item In mProviders(treatment)" in code
    assert "cboProvider.AddItem CStr(item)" in code
    assert "For Each item In mStandardSlots" in code
    assert "For Each item In mReclinedSlots" in code
    assert "cboTime.AddItem CStr(item)" in code


def test_program_form_exposes_recliner_adjacency_preference():
    code = PATIENT_PROGRAM_FORM_CODE
    assert 'StrComp(treatment, "Ανακλινόμενο", vbTextCompare) = 0' in code
    assert "mReclinedSlots" in code
    assert "το Ανακλινόμενο θα τοποθετείται πριν ή μετά τη ΦΘ" in code


def test_program_form_calls_real_automatic_completion_bridge_for_all_included_treatments():
    runtime = PATIENT_PROGRAM_MODULE_CODE
    form = PATIENT_PROGRAM_FORM_CODE
    assert '"complete-vba "' in runtime
    assert "CompletePatientProgram" in runtime
    assert "BuildAutomaticRequestArguments" in form
    assert "If IsIncluded(treatment) Then" in form
    assert '" --request "' in form
    assert "response = CompletePatientProgram(mPatientId, requestArguments)" in form
    assert "ApplyAutomaticCompletion response" in form


def test_program_form_only_fills_missing_draft_fields_from_solver():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "If Not mDraftDays.Exists(treatment)" in code
    assert "If Not mDraftTimes.Exists(treatment)" in code
    assert "If Not mDraftProviders.Exists(treatment)" in code
    assert 'Case "UNRESOLVED"' in code
    assert "RefreshAllTreatmentRows" in code
    assert "Μπορείς να ανοίξεις οποιαδήποτε θεραπεία και να τις αλλάξεις" in code


def test_save_button_enables_only_when_every_included_treatment_is_complete():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "Private Function TreatmentIsComplete" in code
    assert "Private Function DraftIsComplete" in code
    assert "If RequiresProvider(treatment) Then" in code
    assert "cmdSaveProgram.Enabled = DraftIsComplete()" in code


def test_atomic_save_resolves_all_headers_and_snapshots_before_first_write():
    code = PATIENT_PROGRAM_FORM_CODE
    start = code.index("Private Sub SavePatientProgramAtomic")
    save = code[start:]
    resolve_pos = save.index("timeCols(index) = FindPlannerHeader")
    snapshot_pos = save.index("SnapshotCell ws, patientRow, timeCols(index)")
    events_off_pos = save.index("Application.EnableEvents = False")
    first_write_pos = save.index("ws.Cells(patientRow, timeCols(index)).Value =")
    assert resolve_pos < snapshot_pos < events_off_pos < first_write_pos
    assert 'timeHeader = treatment & "_Ώρα"' in save
    assert 'daysHeader = treatment & "_Ημέρες"' in save
    assert 'ProviderHeaderName = "Ψυχ_Ψυχολόγος"' in code


def test_atomic_save_writes_all_fields_verifies_and_saves_once():
    code = PATIENT_PROGRAM_FORM_CODE
    start = code.index("Private Sub SavePatientProgramAtomic")
    end = code.index("Private Sub cmdSaveProgram_Click", start)
    save = code[start:end]
    assert 'NumberFormat = "hh:mm"' in save
    assert "Αποτυχία επαλήθευσης ώρας" in save
    assert "Αποτυχία επαλήθευσης ημερών" in save
    assert "Αποτυχία επαλήθευσης provider" in save
    assert save.count("ThisWorkbook.Save") == 1


def test_atomic_save_rolls_back_and_restores_event_state_on_failure():
    code = PATIENT_PROGRAM_FORM_CODE
    start = code.index("Private Sub SavePatientProgramAtomic")
    end = code.index("Private Sub cmdSaveProgram_Click", start)
    save = code[start:end]
    assert "previousEvents = Application.EnableEvents" in save
    assert "RestoreSnapshots ws, oldValues, oldFormats" in save
    assert "If eventsCaptured Then Application.EnableEvents = previousEvents" in save
    assert "Application.EnableEvents = previousEvents" in save


def test_save_button_calls_atomic_program_save():
    code = PATIENT_PROGRAM_FORM_CODE
    assert "Private Sub cmdSaveProgram_Click()" in code
    assert "SavePatientProgramAtomic" in code
    assert "Δεν διατηρήθηκαν μερικές αλλαγές" in code


def test_program_form_stays_modeless_and_guarded_against_double_open():
    code = PATIENT_PROGRAM_MODULE_CODE
    start = code.index("Public Sub OpenPatientScheduling")
    procedure = code[start:]
    assert "If frmPatientScheduling.Visible Then Exit Sub" in procedure
    assert "frmPatientScheduling.Show vbModeless" in procedure
