from __future__ import annotations

from pathlib import Path
import shutil
import sys

from . import patient_registration_vertical_slice as _vertical
from .patient_registration_form_vba import install_patient_form
from .patient_scheduling_form_vba import install_patient_scheduling_ui

PatientRegistrationBridgeVbaError = _vertical.PatientRegistrationBridgeVbaError
_BASE_VERTICAL_FORM_CODE = _vertical._vertical_form_code


def _scheduling_form_code() -> str:
    code = _BASE_VERTICAL_FORM_CODE()
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


def _install_base_patient_form(workbook_path: Path) -> None:
    """Install frmNewPatient on a private prerequisite copy when absent.

    The repository baseline intentionally stays clean.  The vertical bridge
    expects frmNewPatient to exist, so the scheduling preview composes that
    prerequisite on a temporary .xlsm copy instead of mutating the source.
    """

    if sys.platform != "win32":
        raise PatientRegistrationBridgeVbaError(
            "Patient form installation requires Windows with Microsoft Excel"
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
        install_patient_form(
            workbook.VBProject,
            position_control=_position_control,
        )
        workbook.Save()
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Patient form prerequisite installation failed: {exc}"
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
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    prerequisite = output.with_name(f".{output.stem}.patient-form-prerequisite.xlsm")

    if not source.exists():
        raise PatientRegistrationBridgeVbaError(f"Source workbook not found: {source}")
    if source == output:
        raise PatientRegistrationBridgeVbaError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise PatientRegistrationBridgeVbaError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise PatientRegistrationBridgeVbaError(f"Output already exists: {output}")

    prerequisite.unlink(missing_ok=True)
    shutil.copy2(source, prerequisite)

    try:
        _install_base_patient_form(prerequisite)

        original_form_builder = _vertical._vertical_form_code
        try:
            _vertical._vertical_form_code = _scheduling_form_code
            report = _vertical.create_patient_registration_bridge_preview(
                prerequisite,
                output,
                overwrite=overwrite,
            )
        finally:
            _vertical._vertical_form_code = original_form_builder

        _install_scheduling_ui(output)
        return report
    finally:
        prerequisite.unlink(missing_ok=True)
