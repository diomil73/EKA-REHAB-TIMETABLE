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


def test_day_amendment_action_prepares_before_handoff_and_closes_only_after_worker_launch():
    source = DAY_AMENDMENT_MODULE_CODE

    assert "prepare_day_amendment.py" in source
    assert "day_amendment_worker.py" in source
    assert "ThisWorkbook.Save" in source
    assert 'CreateObject("WScript.Shell").Run(commandLine, 0, True)' in source
    assert 'CreateObject("WScript.Shell").Run commandLine, 0, False' in source
    assert "ThisWorkbook.Close SaveChanges:=False" in source

    prepare_pos = source.index("prepare_day_amendment.py")
    worker_pos = source.index("day_amendment_worker.py")
    close_pos = source.index("ThisWorkbook.Close SaveChanges:=False")
    assert prepare_pos < worker_pos < close_pos


def test_day_amendment_action_requires_reason_and_summary():
    source = DAY_AMENDMENT_MODULE_CODE

    assert "reasonCode = Trim$(InputBox" in source
    assert "summaryText = Trim$(InputBox" in source
    assert '" --reason-code "' in source
    assert '" --summary "' in source
    assert "If Len(reasonCode) = 0 Then Exit Sub" in source
    assert "If Len(summaryText) = 0 Then Exit Sub" in source


def test_day_amendment_installer_owns_only_its_module_and_master_button():
    source = inspect.getsource(install_day_amendment_action)

    assert "vbproject.VBComponents(DAY_AMENDMENT_MODULE_NAME)" in source
    assert "vbproject.VBComponents.Remove(component)" in source
    assert 'workbook.Worksheets("MASTER_SCHEDULE")' in source
    assert "ws.Shapes(DAY_AMENDMENT_BUTTON_NAME).Delete()" in source
    assert "shape.OnAction = DAY_AMENDMENT_MACRO_NAME" in source
    assert 'shape.TextFrame2.TextRange.Text = "Αποθήκευση ημέρας"' in source
