from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

from .patient_edit_form_vba import PATIENT_EDIT_FORM_NAME, install_patient_edit_form
from .registration_menu_vba import MENU_FORM_NAME, USERFORM_CODE


class PatientEditFlowVbaError(RuntimeError):
    pass


@dataclass(frozen=True)
class PatientEditFlowPreviewReport:
    source_path: str
    output_path: str
    source_unchanged: bool
    edit_form_present: bool
    edit_button_present: bool


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _workbook_has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def _menu_code_with_patient_edit() -> str:
    code = USERFORM_CODE
    code = code.replace('.Height = 350', '.Height = 405', 1)
    code = code.replace(
        '    StyleMenuButton cmdPatient, "Νέος ασθενής", 68\n'
        '    StyleMenuButton cmdTherapist, "Νέος θεραπευτής", 113\n'
        '    StyleMenuButton cmdStudent, "Νέος φοιτητής", 158\n'
        '    StyleMenuButton cmdOutpatientSchedule, "Πρόγραμμα εξωτερικού ασθενή", 203',
        '    StyleMenuButton cmdPatient, "Νέος ασθενής", 68\n'
        '    StyleMenuButton cmdEditPatient, "Επεξεργασία ασθενή", 113\n'
        '    StyleMenuButton cmdTherapist, "Νέος θεραπευτής", 158\n'
        '    StyleMenuButton cmdStudent, "Νέος φοιτητής", 203\n'
        '    StyleMenuButton cmdOutpatientSchedule, "Πρόγραμμα εξωτερικού ασθενή", 248',
        1,
    )
    code = code.replace('.Top = 260', '.Top = 305', 1)
    marker = '''Private Sub cmdPatient_Click()
    frmNewPatient.Show
End Sub
'''
    replacement = marker + '''
Private Sub cmdEditPatient_Click()
    frmEditPatient.Show
End Sub
'''
    if marker not in code:
        raise RuntimeError("Patient menu click marker missing")
    return code.replace(marker, replacement, 1)


def create_patient_edit_flow_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> PatientEditFlowPreviewReport:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise PatientEditFlowVbaError(f"Source workbook not found: {source}")
    if source == output:
        raise PatientEditFlowVbaError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise PatientEditFlowVbaError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise PatientEditFlowVbaError(f"Output already exists: {output}")
    if sys.platform != "win32":
        raise PatientEditFlowVbaError("Patient edit flow installation requires Windows with Microsoft Excel")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientEditFlowVbaError("pywin32 is required") from exc

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    excel = None
    workbook = None
    edit_form_present = False
    edit_button_present = False
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject

        install_patient_edit_form(
            vbproject,
            position_control=lambda control, left, top: (
                setattr(control, "Left", left),
                setattr(control, "Top", top),
            ),
        )

        menu = vbproject.VBComponents(MENU_FORM_NAME)
        designer = menu.Designer
        try:
            designer.Controls("cmdEditPatient")
        except Exception:
            button = designer.Controls.Add("Forms.CommandButton.1", "cmdEditPatient", True)
            button.Caption = "Επεξεργασία ασθενή"
            button.Left = 24
            button.Top = 92

        code_module = menu.CodeModule
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(_menu_code_with_patient_edit())

        workbook.Save()
        edit_form_present = vbproject.VBComponents(PATIENT_EDIT_FORM_NAME) is not None
        edit_button_present = designer.Controls("cmdEditPatient") is not None
    except Exception as exc:
        raise PatientEditFlowVbaError(f"Patient edit flow installation failed: {exc}") from exc
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

    if _sha256(source) != source_before:
        output.unlink(missing_ok=True)
        raise PatientEditFlowVbaError("Source workbook changed during patient edit flow installation")
    if not _workbook_has_vba(output) or not edit_form_present or not edit_button_present:
        output.unlink(missing_ok=True)
        raise PatientEditFlowVbaError("Patient edit flow verification failed")

    return PatientEditFlowPreviewReport(
        source_path=str(source),
        output_path=str(output),
        source_unchanged=True,
        edit_form_present=True,
        edit_button_present=True,
    )
