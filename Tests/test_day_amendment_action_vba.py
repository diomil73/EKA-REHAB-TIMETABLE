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


def test_day_amendment_action_prepares_before_handoff_and_uses_proven_app_exit_flow():
    source = DAY_AMENDMENT_MODULE_CODE

    assert "prepare_day_amendment.py" in source
    assert "day_amendment_worker.py" in source
    assert "ThisWorkbook.Save" in source
    assert 'CreateObject("WScript.Shell").Run(commandLine, 0, True)' in source
    assert 'CreateObject("WScript.Shell").Run commandLine, 0, False' in source
    assert "CloseForDaySaveHandoff" in source

    prepare_pos = source.index("prepare_day_amendment.py")
    worker_pos = source.index("day_amendment_worker.py")
    handoff_pos = source.index("CloseForDaySaveHandoff")
    assert prepare_pos < worker_pos < handoff_pos

    # Mirror the already-working Save & Exit lifecycle: restore UI, suppress
    # workbook events during close, quit Excel when this is the only workbook,
    # otherwise close just this workbook and restore the prior event state.
    assert "RestoreExcelInterface" in source
    assert "eventsWereEnabled = Application.EnableEvents" in source
    assert "workbookCount = Application.Workbooks.Count" in source
    assert "Application.EnableEvents = False" in source
    assert "If workbookCount <= 1 Then" in source
    assert "Application.Quit" in source
    assert "ThisWorkbook.Close SaveChanges:=False" in source
    assert "Application.EnableEvents = eventsWereEnabled" in source


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
