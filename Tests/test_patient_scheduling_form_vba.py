import inspect

from rehab_excel.patient_registration_scheduling_slice import _scheduling_form_code
from rehab_excel.patient_scheduling_form_vba import (
    PATIENT_SCHEDULING_FORM_CODE,
    PATIENT_SCHEDULING_MODULE_CODE,
    _ensure_control,
)


def test_registration_continues_directly_into_patient_scheduling():
    code = _scheduling_form_code()
    assert "OpenPatientScheduling patientId" in code
    assert "Συνεχίζουμε στον προγραμματισμό θεραπείας." in code
    assert "ThisWorkbook.Close" not in code
    assert "Application.Quit" not in code


def test_registration_hides_parent_before_opening_scheduling_then_unloads_it():
    code = _scheduling_form_code()
    hide_pos = code.index("Me.Hide")
    open_pos = code.index("OpenPatientScheduling patientId")
    unload_pos = code.index("Unload Me", open_pos)
    assert hide_pos < open_pos < unload_pos


def test_scheduling_form_is_modeless_and_guarded_against_double_open():
    code = PATIENT_SCHEDULING_MODULE_CODE
    start = code.index("Public Sub OpenPatientScheduling")
    procedure = code[start:]
    assert "If frmPatientScheduling.Visible Then Exit Sub" in procedure
    assert "frmPatientScheduling.Show vbModeless" in procedure
    assert "frmPatientScheduling.Show\n" not in procedure


def test_registration_success_is_not_reclassified_when_scheduling_fails():
    code = _scheduling_form_code()
    assert "On Error GoTo SchedulingError" in code
    assert "SchedulingError:" in code
    assert "Ο ασθενής έχει ήδη καταχωρηθεί επιτυχώς." in code
    assert "Δεν μπόρεσε να ανοίξει/ολοκληρωθεί ο προγραμματισμός θεραπείας" in code


def test_scheduling_form_loads_treatments_day_patterns_and_workload_candidates():
    code = PATIENT_SCHEDULING_FORM_CODE
    assert 'valueText = "ΦΘ" Or valueText = "Ρομποτικό"' in code
    assert "cboDayPattern.AddItem valueText" in code
    assert "LoadRecurringSchedulingOptions" in code
    assert 'fields(0) = "THERAPIST"' in code
    assert "cboTherapist.AddItem fields(3)" in code


def test_scheduling_header_shows_patient_name_and_id():
    form_code = PATIENT_SCHEDULING_FORM_CODE
    runtime_code = PATIENT_SCHEDULING_MODULE_CODE
    assert "FindPatientDisplayName(mPatientId)" in form_code
    assert 'lblPatient.Caption = "Ασθενής: " & patientName & "   |   Patient ID: " & mPatientId' in form_code
    assert "Public Function FindPatientDisplayName" in runtime_code
    assert "ws.Cells(rowIndex, 3).Value" in runtime_code


def test_scheduling_form_uses_numbered_human_facing_labels():
    code = PATIENT_SCHEDULING_FORM_CODE
    assert 'lblTitle.Caption = "Πρόγραμμα θεραπείας ασθενούς"' in code
    assert "Επίλεξε θεραπεία και ημέρες για να εμφανιστούν οι διαθέσιμες επιλογές." in code
    installer = __import__(
        "rehab_excel.patient_scheduling_form_vba",
        fromlist=["install_patient_scheduling_ui"],
    )
    source = installer.install_patient_scheduling_ui.__code__.co_consts
    joined = " ".join(str(item) for item in source)
    assert "1. Θεραπεία" in joined
    assert "2. Ημέρες θεραπείας" in joined
    assert "3. Διαθέσιμος θεραπευτής" in joined
    assert "4. Διαθέσιμη ώρα" in joined


def test_scheduling_controls_receive_non_default_widths():
    source = inspect.getsource(_ensure_control)
    assert "width = 430" in source
    assert "width = 260" in source
    assert "width = 220" in source
    assert "control.Width = width" in source
    assert "control.Height = height" in source


def test_scheduling_form_only_shows_slots_for_selected_therapist():
    code = PATIENT_SCHEDULING_FORM_CODE
    assert "mSlotsByTherapist(fields(2)) = fields(4)" in code
    assert 'slots = Split(CStr(mSlotsByTherapist(therapistId)), "|")' in code
    assert "cboTime.AddItem" in code


def test_auto_proposal_prefills_controls_without_saving():
    code = PATIENT_SCHEDULING_FORM_CODE
    start = code.index("Private Sub cmdAutoPropose_Click()")
    end = code.index("Private Sub cmdSave_Click()", start)
    procedure = code[start:end]
    assert "ProposeAutomaticSchedule(" in procedure
    assert 'fields(0) = "ASSIGNMENT"' in procedure
    assert "cboTherapist.ListIndex = therapistIndex" in procedure
    assert "cboTime.Value = proposedTime" in procedure
    assert "πάτησε Αποθήκευση για επιβεβαίωση" in procedure
    assert "SavePatientRecurringAssignment" not in procedure


def test_runtime_module_exposes_auto_proposal_bridge_and_preview_path():
    code = PATIENT_SCHEDULING_MODULE_CODE
    assert 'args = "auto-propose-vba "' in code
    assert '" --request " & QuoteArg(Trim$(treatment) & "|" & Trim$(dayPattern))' in code
    assert 'ThisWorkbook.Path & "\\..\\..\\Python\\tools\\patient_scheduling_bridge.py"' in code


def test_multiline_function_calls_use_vba_line_continuations():
    code = PATIENT_SCHEDULING_FORM_CODE
    assert "response = LoadRecurringSchedulingOptions( _\n" in code
    assert "response = ProposeAutomaticSchedule( _\n" in code
    assert "response = ValidateRecurringSelection( _\n" in code
    assert "        mPatientId, _\n" in code
    assert "        Trim$(cboDayPattern.Value) _\n    )" in code
    assert "        Trim$(cboTime.Value) _\n    )" in code


def test_save_revalidates_before_authoritative_planner_write():
    code = PATIENT_SCHEDULING_FORM_CODE
    validate_pos = code.index("ValidateRecurringSelection(")
    save_pos = code.index("SavePatientRecurringAssignment _")
    assert validate_pos < save_pos
    assert "LoadSelections" in code[validate_pos:save_pos]


def test_runtime_module_writes_supported_recurring_planner_columns_and_saves():
    code = PATIENT_SCHEDULING_MODULE_CODE
    assert 'Case "ΦΘ": prefix = "ΦΘ"' in code
    assert 'Case "Ρομποτικό": prefix = "Ρομποτικό"' in code
    assert 'FindPlannerHeader(ws, prefix & "_Ώρα")' in code
    assert 'FindPlannerHeader(ws, prefix & "_Ημέρες")' in code
    assert 'FindPlannerHeader(ws, prefix & "_Θεραπευτής")' in code
    assert "ThisWorkbook.Save" in code


def test_runtime_bridge_is_hidden_and_waits_for_fresh_result():
    code = PATIENT_SCHEDULING_MODULE_CODE
    assert 'CreateObject("WScript.Shell")' in code
    assert "shell.Run(command, 0, True)" in code
    assert "PYTHONIOENCODING=utf-8" in code
    assert 'Charset = "utf-8"' in code
