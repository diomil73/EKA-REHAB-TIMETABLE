from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Protocol
from zipfile import ZipFile

from .application_shell_vba import APP_SHELL_MODULE_NAME, install_application_shell
from .daily_input_action_vba import DAILY_INPUT_MODULE_NAME, install_daily_input_action
from .master_toolbar_vba import (
    MASTER_TOOLBAR_MODULE_NAME,
    MasterToolbarError,
    install_master_toolbar,
)
from .patient_registration_form_vba import PATIENT_FORM_NAME, install_patient_form
from .therapist_registration_form_vba import THERAPIST_FORM_NAME, install_therapist_form
from .student_registration_form_vba import STUDENT_FORM_NAME, install_student_form


MENU_FORM_NAME = "frmRegistrationMenu"
MENU_MODULE_NAME = "modRegistrationMenu"
MENU_MACRO_NAME = "ShowRegistrationMenu"
DEFAULT_PREVIEW_FILENAME = "REGISTRATION_MENU_PREVIEW.xlsm"


STANDARD_MODULE_CODE = f'''Option Explicit

Public Sub {MENU_MACRO_NAME}()
    {MENU_FORM_NAME}.Show
End Sub
'''


USERFORM_CODE = '''Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Κεντρικό Μενού"
        .Width = 330
        .Height = 400
        .StartUpPosition = 1
        .BackColor = RGB(245, 247, 250)
    End With

    With lblTitle
        .Caption = "Επιλέξτε ενέργεια"
        .Left = 35
        .Top = 24
        .Width = 250
        .Height = 26
        .TextAlign = 2
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(45, 55, 72)
        .BackStyle = 0
    End With

    StyleMenuButton cmdPatient, "Νέος ασθενής", 68
    StyleMenuButton cmdTherapist, "Νέος θεραπευτής", 113
    StyleMenuButton cmdStudent, "Νέος φοιτητής", 158
    StyleMenuButton cmdOutpatientSchedule, "Πρόγραμμα εξωτερικού ασθενή", 203
    StyleMenuButton cmdDailyInput, "Εφαρμογή DAILY_INPUT", 248

    With cmdClose
        .Caption = "Κλείσιμο"
        .Left = 105
        .Top = 310
        .Width = 110
        .Height = 30
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Cancel = True
    End With
End Sub

Private Sub StyleMenuButton(ByVal button As MSForms.CommandButton, ByVal text As String, ByVal topPosition As Single)
    With button
        .Caption = text
        .Left = 45
        .Top = topPosition
        .Width = 230
        .Height = 34
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub cmdPatient_Click()
    frmNewPatient.Show
End Sub

Private Sub cmdTherapist_Click()
    frmNewTherapist.Show
End Sub

Private Sub cmdStudent_Click()
    frmNewStudent.Show
End Sub

Private Sub cmdOutpatientSchedule_Click()
    On Error GoTo MissingForm
    frmOutpatientSchedule.Show
    Exit Sub
MissingForm:
    MsgBox "Η φόρμα προγράμματος εξωτερικού ασθενή δεν είναι εγκατεστημένη σε αυτό το αρχείο.", vbExclamation, "Πρόγραμμα εξωτερικού ασθενή"
End Sub

Private Sub cmdDailyInput_Click()
    ApplyDailyInputPreview
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'''


class RegistrationMenuVbaError(RuntimeError):
    """Raised when a safe registration-menu preview cannot be produced."""


@dataclass(frozen=True)
class RegistrationMenuPreviewReport:
    source_path: str
    output_path: str
    source_unchanged: bool
    vba_present: bool
    module_present: bool
    form_present: bool
    macro_name: str = MENU_MACRO_NAME


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def workbook_has_vba(path: str | Path) -> bool:
    workbook_path = Path(path)
    try:
        with ZipFile(workbook_path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


class RegistrationMenuInstallerBackend(Protocol):
    def install(self, workbook_path: Path) -> tuple[bool, bool]:
        """Install menu VBA and return (module_present, form_present)."""
        ...


class Win32ComRegistrationMenuInstaller:
    """Install the central registration menu into an already-created workbook copy."""

    @staticmethod
    def _remove_component_if_present(vbproject, name: str) -> None:
        try:
            component = vbproject.VBComponents(name)
        except Exception:
            return
        vbproject.VBComponents.Remove(component)

    @staticmethod
    def _position_control(control, left: float, top: float) -> None:
        control.Left = left
        control.Top = top

    @classmethod
    def _add_command_button(cls, designer, name: str, caption: str, top: float) -> None:
        button = designer.Controls.Add("Forms.CommandButton.1", name, True)
        button.Caption = caption
        cls._position_control(button, 24, top)

    @classmethod
    def _add_label(cls, designer, name: str, caption: str, top: float) -> None:
        label = designer.Controls.Add("Forms.Label.1", name, True)
        label.Caption = caption
        cls._position_control(label, 24, top)

    @staticmethod
    def _component_present(vbproject, name: str) -> bool:
        try:
            _ = vbproject.VBComponents(name)
            return True
        except Exception:
            return False

    def install(self, workbook_path: Path) -> tuple[bool, bool]:
        if sys.platform != "win32":
            raise RegistrationMenuVbaError(
                "Registration menu installation requires Windows with Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise RegistrationMenuVbaError(
                "pywin32 is required for registration menu installation"
            ) from exc

        excel = None
        workbook = None
        stage = "starting Excel"
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            stage = "opening workbook"
            workbook = excel.Workbooks.Open(
                str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False
            )

            stage = "accessing VBA project"
            try:
                vbproject = workbook.VBProject
                _ = vbproject.VBComponents.Count
            except Exception as exc:
                raise RegistrationMenuVbaError(
                    "Excel blocked programmatic access to the VBA project. "
                    "Enable File > Options > Trust Center > Trust Center Settings > "
                    "Macro Settings > Trust access to the VBA project object model, "
                    "then retry on the preview copy."
                ) from exc

            stage = "removing old menu module"
            self._remove_component_if_present(vbproject, MENU_MODULE_NAME)
            stage = "removing old menu form"
            self._remove_component_if_present(vbproject, MENU_FORM_NAME)
            stage = "removing old app shell module"
            self._remove_component_if_present(vbproject, APP_SHELL_MODULE_NAME)
            stage = "removing old MASTER toolbar module"
            self._remove_component_if_present(vbproject, MASTER_TOOLBAR_MODULE_NAME)
            stage = "removing old DAILY_INPUT action module"
            self._remove_component_if_present(vbproject, DAILY_INPUT_MODULE_NAME)
            stage = "removing old patient form"
            self._remove_component_if_present(vbproject, PATIENT_FORM_NAME)
            stage = "removing old therapist form"
            self._remove_component_if_present(vbproject, THERAPIST_FORM_NAME)
            stage = "removing old student form"
            self._remove_component_if_present(vbproject, STUDENT_FORM_NAME)

            stage = "creating menu module"
            module = vbproject.VBComponents.Add(1)
            module.Name = MENU_MODULE_NAME
            module.CodeModule.AddFromString(STANDARD_MODULE_CODE)

            stage = "adding blank menu UserForm"
            form = vbproject.VBComponents.Add(3)
            stage = "naming menu UserForm"
            form.Name = MENU_FORM_NAME
            stage = "accessing menu designer"
            designer = form.Designer
            stage = "setting menu caption"
            designer.Caption = "Κεντρικό Μενού"

            stage = "adding menu title"
            self._add_label(designer, "lblTitle", "Επιλέξτε ενέργεια", 18)
            stage = "adding patient menu button"
            self._add_command_button(designer, "cmdPatient", "Νέος ασθενής", 55)
            stage = "adding therapist menu button"
            self._add_command_button(designer, "cmdTherapist", "Νέος θεραπευτής", 92)
            stage = "adding student menu button"
            self._add_command_button(designer, "cmdStudent", "Νέος φοιτητής", 129)
            stage = "adding outpatient schedule menu button"
            self._add_command_button(
                designer,
                "cmdOutpatientSchedule",
                "Πρόγραμμα εξωτερικού ασθενή",
                166,
            )
            stage = "adding DAILY_INPUT menu button"
            self._add_command_button(
                designer,
                "cmdDailyInput",
                "Εφαρμογή DAILY_INPUT",
                203,
            )
            stage = "adding close menu button"
            self._add_command_button(designer, "cmdClose", "Κλείσιμο", 250)

            stage = "writing menu form code"
            form.CodeModule.AddFromString(USERFORM_CODE)

            stage = "installing application shell"
            install_application_shell(vbproject, workbook=workbook)
            stage = "installing MASTER toolbar"
            install_master_toolbar(vbproject, workbook)
            stage = "installing DAILY_INPUT action"
            install_daily_input_action(vbproject)
            stage = "installing patient form"
            install_patient_form(vbproject, position_control=self._position_control)
            stage = "installing therapist form"
            install_therapist_form(vbproject, position_control=self._position_control)
            stage = "installing student form"
            install_student_form(vbproject, position_control=self._position_control)

            stage = "saving workbook"
            workbook.Save()

            stage = "verifying menu components"
            module_present = self._component_present(vbproject, MENU_MODULE_NAME)
            form_present = self._component_present(vbproject, MENU_FORM_NAME)
            patient_form_present = self._component_present(vbproject, PATIENT_FORM_NAME)
            therapist_form_present = self._component_present(vbproject, THERAPIST_FORM_NAME)
            student_form_present = self._component_present(vbproject, STUDENT_FORM_NAME)
            app_shell_present = self._component_present(vbproject, APP_SHELL_MODULE_NAME)
            master_toolbar_present = self._component_present(vbproject, MASTER_TOOLBAR_MODULE_NAME)
            daily_input_action_present = self._component_present(vbproject, DAILY_INPUT_MODULE_NAME)
            if not app_shell_present:
                raise RegistrationMenuVbaError(
                    "Application shell module was not confirmed in the preview VBA project"
                )
            if not master_toolbar_present:
                raise RegistrationMenuVbaError(
                    "MASTER toolbar module was not confirmed in the preview VBA project"
                )
            if not patient_form_present:
                raise RegistrationMenuVbaError(
                    "Patient registration form was not confirmed in the preview VBA project"
                )
            if not therapist_form_present:
                raise RegistrationMenuVbaError(
                    "Therapist registration form was not confirmed in the preview VBA project"
                )
            if not student_form_present:
                raise RegistrationMenuVbaError(
                    "Student registration form was not confirmed in the preview VBA project"
                )
            if not daily_input_action_present:
                raise RegistrationMenuVbaError(
                    "DAILY_INPUT action module was not confirmed in the preview VBA project"
                )
            return module_present, form_present
        except (RegistrationMenuVbaError, MasterToolbarError):
            raise
        except Exception as exc:
            raise RegistrationMenuVbaError(
                f"Excel registration menu installation failed during {stage}: {exc}"
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


def create_registration_menu_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    backend: RegistrationMenuInstallerBackend | None = None,
    overwrite: bool = False,
) -> RegistrationMenuPreviewReport:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise RegistrationMenuVbaError(f"Source workbook not found: {source}")
    if source == output:
        raise RegistrationMenuVbaError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise RegistrationMenuVbaError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise RegistrationMenuVbaError(f"Output already exists: {output}")

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        try:
            output.unlink()
        except PermissionError as exc:
            raise RegistrationMenuVbaError(
                f"Preview workbook is open or locked by Excel: {output}"
            ) from exc

    shutil.copy2(source, output)
    selected_backend = backend or Win32ComRegistrationMenuInstaller()

    try:
        module_present, form_present = selected_backend.install(output)
    except Exception:
        try:
            output.unlink(missing_ok=True)
        except PermissionError:
            pass
        raise

    source_after = _sha256(source)
    if source_before != source_after:
        try:
            output.unlink(missing_ok=True)
        except PermissionError:
            pass
        raise RegistrationMenuVbaError("Source workbook changed during menu installation")

    vba_present = workbook_has_vba(output)
    if not vba_present or not module_present or not form_present:
        try:
            output.unlink(missing_ok=True)
        except PermissionError:
            pass
        raise RegistrationMenuVbaError(
            "Menu preview verification failed: expected VBA module/form not confirmed"
        )

    return RegistrationMenuPreviewReport(
        source_path=str(source),
        output_path=str(output),
        source_unchanged=True,
        vba_present=True,
        module_present=True,
        form_present=True,
    )
