import inspect

from rehab_excel.day_amendment_action_vba import (
    DAY_AMENDMENT_BUTTON_NAME,
    DAY_AMENDMENT_MACRO_NAME,
    DAY_AMENDMENT_MODULE_CODE,
    DAY_AMENDMENT_MODULE_NAME,
    install_day_amendment_action,
)


def test_day_amendment_action_is_explicit_and_separate_from_save_exit():
    source = DAY_AMENDMENT_MODULE_CODE

    assert DAY_AMENDMENT_MODULE_NAME == "modDayAmendmentAction"
    assert DAY_AMENDMENT_MACRO_NAME == "StartClosedDayAmendment"
    assert DAY_AMENDMENT_BUTTON_NAME == "ekaToolbar_DayAmendment"
    assert "Αποθήκευση ημέρας" in source
    assert "vbYesNo + vbDefaultButton2" in source
    assert "Η προηγούμενη καταχώρηση θα παραμείνει στο ιστορικό" in source
    assert "ExitApplication" not in source
    assert "Save & Exit" not in source


def test_day_save_is_fully_in_session_without_close_worker_or_reopen():
    source = DAY_AMENDMENT_MODULE_CODE

    assert "prepare_day_amendment_in_session.py" in source
    assert "PersistDaySavePayload(payloadPath)" in source
    assert "ThisWorkbook.Save" in source
    assert 'CreateObject("WScript.Shell").Run(commandLine, 0, True)' in source

    assert "day_amendment_worker.py" not in source
    assert "CloseForDaySaveHandoff" not in source
    assert "Application.Quit" not in source
    assert "ThisWorkbook.Close" not in source
    assert "os.startfile" not in source
    assert 'Run commandLine, 0, False' not in source


def test_day_save_in_session_persistence_is_preflight_readback_save_and_rollback():
    source = DAY_AMENDMENT_MODULE_CODE

    assert 'Case "SUPERSEDE"' in source
    assert 'Case "SNAPSHOT"' in source
    assert 'Case "THERAPIST"' in source
    assert 'Case "PATIENT"' in source
    assert 'Case "EVENT"' in source
    assert "oldSnapshotRow = FindRowById" in source
    assert '"CURRENT_FINAL"' in source
    assert "HistoryPairExists" in source
    assert "PatientHistoryPairExists" in source
    assert "SupersededSnapshotMatches" in source
    assert "RowMatchesPayload" in source
    assert "CellMatchesPayload" in source
    assert "oldStatus = CStr" in source
    assert "wsSnapshot.Cells(oldSnapshotRow, 4).Value2 = oldStatus" in source
    assert "ClearContents" in source
    assert "FinalizeTechnicalSheet" in source
    assert "ThisWorkbook.Save" in source


def test_day_save_patient_history_is_required_duplicate_safe_and_atomic():
    source = DAY_AMENDMENT_MODULE_CODE

    assert 'Private Const PATIENT_SHEET As String = "PATIENT_SESSION_HISTORY"' in source
    assert "If UBound(parts) <> 14" in source
    assert "Set wsPatient = ExistingSheet(PATIENT_SHEET)" in source
    assert "wsSnapshot Is Nothing Or wsTherapist Is Nothing Or wsPatient Is Nothing Or wsEvent Is Nothing" in source
    assert "If Not HeadersMatch(wsPatient, PatientHeaders()) Then" in source
    assert "PatientHistoryPairExists(wsPatient, PayloadValue(parts(1)), PayloadValue(parts(4)))" in source
    assert "patientStartRow = NextWriteRow(wsPatient)" in source
    assert "WritePayloadFields wsPatient, patientStartRow + i - 1, parts, 1, 14" in source
    assert "RowMatchesPayload(wsPatient, patientStartRow + i - 1, parts, 1, 14)" in source
    assert 'wsPatient.Rows(CStr(patientStartRow) & ":" & CStr(patientStartRow + patientCount - 1)).ClearContents' in source
    assert "FinalizeTechnicalSheet wsPatient" in source


def test_supersede_readback_checks_identity_version_and_status_without_full_row_serialization():
    source = DAY_AMENDMENT_MODULE_CODE

    assert "Private Function SupersededSnapshotMatches" in source
    assert "CellMatchesPayload(ws.Cells(targetRow, 1).Value, PayloadValue(parts(1)))" in source
    assert "CellMatchesPayload(ws.Cells(targetRow, 3).Value, PayloadValue(parts(3)))" in source
    assert "PayloadValue(parts(4))" in source
    assert "SupersededSnapshotMatches(wsSnapshot, oldSnapshotRow, supersedeParts)" in source
    assert "RowMatchesPayload(wsSnapshot, oldSnapshotRow, supersedeParts, 1, 15)" not in source


def test_row_readback_normalizes_numeric_iso_date_and_time_excel_values():
    source = DAY_AMENDMENT_MODULE_CODE

    assert "Private Function CellMatchesPayload" in source
    assert "If IsNumeric(actualValue) And IsNumeric(expectedText) Then" in source
    assert "CDbl(actualValue) = CDbl(expectedText)" in source
    assert "If IsIsoDateText(expectedText) And IsDate(actualValue) Then" in source
    assert 'Format$(CDate(actualValue), "yyyy-mm-dd") = expectedText' in source
    assert "Private Function IsIsoDateText" in source
    assert "If IsIsoTimeText(expectedText) And IsDate(actualValue) Then" in source
    assert 'Format$(CDate(actualValue), "hh:nn:ss") = expectedText' in source
    assert "Private Function IsIsoTimeText" in source
    assert "If Not CellMatchesPayload(ws.Cells(targetRow, i + 1).Value" in source
    assert 'PatientHeaders = Array("SnapshotId", "ServiceDate", "SnapshotVersion", "SessionId", "PatientId", "Status", "OriginalTherapistId", "OriginalTime", "EffectiveTherapistId", "EffectiveTime", "ReplacementId", "CancellationId", "CancellationKind", "Reason")' in source


def test_day_amendment_action_has_no_manual_reason_or_summary_prompts():
    source = DAY_AMENDMENT_MODULE_CODE

    assert "InputBox(" not in source
    assert "reasonCode" not in source
    assert "summaryText" not in source
    assert '" --reason-code "' not in source
    assert '" --summary "' not in source
    assert '" --actor-or-source "' in source


def test_day_save_installer_touches_only_daily_input_not_master():
    source = inspect.getsource(install_day_amendment_action)

    assert "vbproject.VBComponents(DAY_AMENDMENT_MODULE_NAME)" in source
    assert "vbproject.VBComponents.Remove(component)" in source
    assert 'workbook.Worksheets("MASTER_SCHEDULE")' not in source
    assert 'workbook.Worksheets("DAILY_INPUT")' in source
    assert "last_row = _last_used_row(ws)" in source
    assert 'anchor = ws.Range(f"A{last_row + 2}")' in source
    assert "shape.OnAction = DAY_AMENDMENT_MACRO_NAME" in source
    assert 'shape.TextFrame2.TextRange.Text = "Αποθήκευση ημέρας"' in source
