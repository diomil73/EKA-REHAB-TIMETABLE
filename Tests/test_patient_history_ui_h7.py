from pathlib import Path

from rehab_excel.patient_history_report_vba import (
    PATIENT_REPORT_MODULE,
    PATIENT_REPORT_NAV,
    PATIENT_REPORT_OPEN,
    PATIENT_REPORT_SHEET,
    PATIENT_REPORT_VBA,
)
from tools import install_patient_history_ui as installer


def test_patient_report_four_modes_and_existing_h5_export():
    src = PATIENT_REPORT_VBA
    assert PATIENT_REPORT_MODULE == "modPatientHistoryReport"
    assert PATIENT_REPORT_SHEET == "PATIENT_HISTORY_REPORT"
    assert PATIENT_REPORT_OPEN == "OpenPatientHistoryReport"
    assert 'RenderPatientHistory "daily"' in src
    assert 'RenderPatientHistory "weekly"' in src
    assert 'RenderPatientHistory "monthly"' in src
    assert 'RenderPatientHistory "overall"' in src
    assert "export_patient_history_ui.py" in src
    assert 'CreateObject("WScript.Shell").Run(commandLine, 0, True)' in src
    assert "ThisWorkbook.Close" not in src
    assert "Application.Quit" not in src


def test_patient_report_read_only_history_and_hidden_snapshot_version():
    src = PATIENT_REPORT_VBA
    assert "FINALIZED_DAY_SNAPSHOTS" not in src
    assert "PATIENT_SESSION_HISTORY" not in src
    assert '"Έκδοση"' not in src
    assert 'ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True' in src
    assert 'If UBound(fields) <> 8' in src


def test_report_buttons_persist_through_mode_switches():
    src = PATIENT_REPORT_VBA
    for mode in ("Daily", "Weekly", "Monthly", "Overall", "Back"):
        assert f'EnsurePatientButton ws, "{mode}"' in src
    assert "PositionPatientButtons ws" in src
    assert "shape.Visible = True" in src
    assert "shape.ZOrder 0" in src
    assert "ResetPatientReportWindow ws" in src
    assert "ActiveWindow.FreezePanes = False" in src


def test_error_restores_excel_session_state():
    src = PATIENT_REPORT_VBA
    assert src.index("priorEvents = Application.EnableEvents") < src.index("On Error GoTo RenderError")
    assert "Application.EnableEvents = priorEvents" in src
    assert "Application.ScreenUpdating = priorScreenUpdating" in src
    assert "If Not previousSheet Is Nothing Then previousSheet.Activate" in src
    assert "Resume RenderCleanup" in src


def test_native_installer_persisted_reopen_and_history_guard():
    code = Path(installer.__file__).read_text(encoding="utf-8")
    assert PATIENT_REPORT_NAV == "ekaToolbar_PatientHistory"
    assert "ReadOnly=True" in code
    assert "Persisted VBA reopen verification: OK" in code
    for sheet in ("FINALIZED_DAY_SNAPSHOTS", "THERAPIST_DAY_HISTORY", "PATIENT_SESSION_HISTORY", "DAY_HISTORY_EVENTS"):
        assert sheet in code
    assert "int(workbook.Worksheets(sheet_name).Visible) != 2" in code

def test_patient_report_installer_reprotects_blank_report_and_keeps_shapes_clickable():
    import inspect

    from rehab_excel.patient_history_report_vba import install_patient_report_components

    source = inspect.getsource(install_patient_report_components)
    assert 'report.Protect("", False, True, True)' in source
    assert "ProtectContents={protect_contents}" in source
    assert "ProtectDrawingObjects={protect_objects}" in source


def test_patient_report_native_audit_checks_persisted_protection():
    import inspect

    source = inspect.getsource(installer._audit)
    assert "report.ProtectContents" in source
    assert "report.ProtectDrawingObjects" in source
    assert "Patient report protection state is invalid" in source

