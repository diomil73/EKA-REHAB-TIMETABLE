from __future__ import annotations

from pathlib import Path
import sys

from . import patient_registration_vertical_slice as _vertical
from .patient_scheduling_form_vba import install_patient_scheduling_ui

PatientRegistrationBridgeVbaError = _vertical.PatientRegistrationBridgeVbaError


def _scheduling_form_code() -> str:
    code = _vertical._vertical_form_code()
    old = '''    Unload Me
    GoToMaster
    MsgBox "Ο ασθενής καταχωρήθηκε επιτυχώς." & vbCrLf & _
           "Patient ID: " & patientId, _
           vbInformation, "Νέος ασθενής"
    Exit Sub'''
    new = '''    MsgBox "Ο ασθενής καταχωρήθηκε επιτυχώς." & vbCrLf & _
           "Patient ID: " & patientId & vbCrLf & vbCrLf & _
           "Συνεχίζουμε στον προγραμματισμό θεραπείας.", _
           vbInformation, "Νέος ασθενής"
    Unload Me
    OpenPatientScheduling patientId
    GoToMaster
    Exit Sub'''
    if old not in code:
        raise RuntimeError("Patient scheduling continuation marker missing")
    return code.replace(old, new, 1)


def _position_control(control, left: float, top: float) -> None:
    try:
        control.Left = left
    except Exception:
        pass
    try:
        control.Top = top
    except Exception:
        pass


def _install_scheduling_ui(workbook_path: Path) -> None:
    if sys.platform != "win32":
        raise PatientRegistrationBridgeVbaError(
            "Patient scheduling UI installation requires Windows with Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientRegistrationBridgeVbaError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)
        install_patient_scheduling_ui(
            workbook.VBProject,
            position_control=_position_control,
        )
        workbook.Save()
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Patient scheduling UI installation failed: {exc}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def create_patient_registration_bridge_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
):
    original_form_builder = _vertical._vertical_form_code
    try:
        _vertical._vertical_form_code = _scheduling_form_code
        report = _vertical.create_patient_registration_bridge_preview(
            source_path,
            output_path,
            overwrite=overwrite,
        )
    finally:
        _vertical._vertical_form_code = original_form_builder

    _install_scheduling_ui(Path(output_path))
    return report
