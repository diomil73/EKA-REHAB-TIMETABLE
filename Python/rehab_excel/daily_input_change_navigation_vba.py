from __future__ import annotations


MODULE_NAME = "modDailyInputChangeNavigation"
MACRO_NAME = "OpenTherapistDailyWithChanges"
WORKLOAD_BUTTON_NAME = "ekaWorkload_Refresh"

MODULE_CODE = r'''Option Explicit

Public Sub OpenTherapistDailyWithChanges()
    ShowTherapistWorkload
    On Error Resume Next
    SyncDailyInputChangesToLedger
    KeepApplicationShell
    On Error GoTo 0
End Sub
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def install_daily_input_change_navigation(vbproject, workbook) -> None:
    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)

    shape = workbook.Worksheets("MASTER_SCHEDULE").Shapes(WORKLOAD_BUTTON_NAME)
    shape.OnAction = MACRO_NAME
    shape.Locked = True
