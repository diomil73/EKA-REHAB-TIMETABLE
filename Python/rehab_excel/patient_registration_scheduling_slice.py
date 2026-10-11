from __future__ import annotations

from pathlib import Path
import shutil
import sys
import tempfile

from . import patient_registration_vertical_slice as _vertical
from .native_excel import apply_write_plan_to_copy
from .outpatient_schedule_source import read_unified_base_schedule
from .patient_edit_form_vba import install_patient_edit_form
from .patient_program_combo_clear_vba import install_combo_clear_handlers
from .patient_program_form_vba import install_patient_program_ui
from .patient_program_master_render_vba import install_patient_program_master_render
from .patient_registration_form_vba import PATIENT_FORM_NAME, install_patient_form
from .patient_registry_source import read_patient_registry
from .recurring_daily_style import build_recurring_therapist_daily_style_plan

PatientRegistrationBridgeVbaError = _vertical.PatientRegistrationBridgeVbaError
_BASE_VERTICAL_FORM_CODE = _vertical._vertical_form_code
LEGACY_PATIENT_SCHEDULING_MODULE_NAME = "PatientSchedulingRuntime"
PATIENT_PROGRAM_FORM_NAME = "frmPatientScheduling"
PATIENT_PROGRAM_MODULE_NAME = "PatientProgramRuntime"


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
    On Error GoTo SchedulingError
    Me.Hide
    OpenPatientScheduling patientId
    Unload Me
    GoToMaster
    Exit Sub

SchedulingError:
    On Error Resume Next
    Unload Me
    On Error GoTo 0
    MsgBox "Ο ασθενής έχει ήδη καταχωρηθεί επιτυχώς." & vbCrLf & _
           "Δεν μπόρεσε να ανοίξει/ολοκληρωθεί ο προγραμματισμός θεραπείας:" & vbCrLf & _
           Err.Description, _
           vbExclamation, "Προγραμματισμός θεραπείας"
    On Error Resume Next
    GoToMaster
    On Error GoTo 0
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


def _has_vb_component(vbproject, component_name: str) -> bool:
    """Return True when a named VBA component already exists."""

    try:
        return vbproject.VBComponents(component_name) is not None
    except Exception:
        return False


def _remove_vb_component_if_present(vbproject, component_name: str) -> bool:
    """Remove one obsolete VBA component and report whether it existed."""

    try:
        component = vbproject.VBComponents(component_name)
    except Exception:
        return False
    vbproject.VBComponents.Remove(component)
    return True


def _component_code_text(vbproject, component_name: str) -> str:
    component = vbproject.VBComponents(component_name)
    code_module = component.CodeModule
    count = int(getattr(code_module, "CountOfLines", 0) or 0)
    if count <= 0:
        return ""
    return str(code_module.Lines(1, count))


def _verify_patient_program_vba(vbproject) -> None:
    """Reject previews that still contain the retired single-treatment runtime/form code."""

    if _has_vb_component(vbproject, LEGACY_PATIENT_SCHEDULING_MODULE_NAME):
        raise RuntimeError(
            f"Legacy VBA module still present: {LEGACY_PATIENT_SCHEDULING_MODULE_NAME}"
        )
    if not _has_vb_component(vbproject, PATIENT_PROGRAM_MODULE_NAME):
        raise RuntimeError(f"Missing VBA module: {PATIENT_PROGRAM_MODULE_NAME}")
    if not _has_vb_component(vbproject, PATIENT_PROGRAM_FORM_NAME):
        raise RuntimeError(f"Missing VBA form: {PATIENT_PROGRAM_FORM_NAME}")

    form_code = _component_code_text(vbproject, PATIENT_PROGRAM_FORM_NAME)
    runtime_code = _component_code_text(vbproject, PATIENT_PROGRAM_MODULE_NAME)

    required_form_markers = (
        "lstTreatments",
        "SavePatientProgramAtomic",
        "BuildAutomaticRequestArguments",
    )
    required_runtime_markers = (
        "LoadPatientProgramConfig",
        "CompletePatientProgram",
        "OpenPatientScheduling",
    )
    legacy_markers = (
        "cboTreatment",
        "LoadRecurringSchedulingOptions",
        "SavePatientRecurringAssignment",
    )

    missing_form = [marker for marker in required_form_markers if marker not in form_code]
    missing_runtime = [marker for marker in required_runtime_markers if marker not in runtime_code]
    legacy_found = [
        marker
        for marker in legacy_markers
        if marker in form_code or marker in runtime_code
    ]

    if missing_form or missing_runtime or legacy_found:
        details = []
        if missing_form:
            details.append("missing form markers: " + ", ".join(missing_form))
        if missing_runtime:
            details.append("missing runtime markers: " + ", ".join(missing_runtime))
        if legacy_found:
            details.append("legacy markers still present: " + ", ".join(legacy_found))
        raise RuntimeError("Patient-program VBA verification failed: " + "; ".join(details))


def _install_base_patient_form(workbook_path: Path) -> None:
    """Ensure frmNewPatient exists on a private prerequisite copy."""

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

        if not _has_vb_component(workbook.VBProject, PATIENT_FORM_NAME):
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
    """Install/update the shared patient registration/edit scheduling UI in place."""

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
    stage = "starting Excel"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False

        stage = "opening preview workbook"
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)

        stage = "accessing VBA project"
        vbproject = workbook.VBProject
        _ = vbproject.VBComponents.Count

        stage = "removing legacy patient scheduling runtime"
        _remove_vb_component_if_present(
            vbproject,
            LEGACY_PATIENT_SCHEDULING_MODULE_NAME,
        )

        stage = "installing patient-program form/runtime"
        install_patient_program_ui(
            vbproject,
            position_control=_position_control,
        )

        stage = "installing/updating patient edit form"
        install_patient_edit_form(
            vbproject,
            position_control=_position_control,
        )

        stage = "installing MASTER patient-program renderer"
        install_patient_program_master_render(vbproject)

        stage = "installing combo clear handlers"
        install_combo_clear_handlers(vbproject)

        stage = "verifying patient-program VBA replacement"
        _verify_patient_program_vba(vbproject)

        stage = "saving preview workbook"
        workbook.Save()
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Patient scheduling UI installation failed during {stage}: {exc}"
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


def _refresh_daily_visuals(workbook_path: Path) -> None:
    """Refresh recurring THERAPIST_DAILY rich-text roles from authoritative data."""

    try:
        patients = read_patient_registry(workbook_path)
        base_entries = read_unified_base_schedule(workbook_path)
        write_plan = build_recurring_therapist_daily_style_plan(
            workbook_path,
            base_entries,
            patients,
        )
        if write_plan.operation_count == 0:
            return

        # Keep the intermediate rich-text workbook out of Excel/previews. On
        # Windows, a leading dot does not hide files, so the old sibling temp
        # copy was visible to users as e.g. V53.recurring-style.xlsm.
        with tempfile.TemporaryDirectory(prefix="eka-recurring-style-") as temp_dir:
            styled_copy = Path(temp_dir) / f"{workbook_path.stem}.xlsm"
            apply_write_plan_to_copy(write_plan, styled_copy, overwrite=True)
            shutil.copy2(styled_copy, workbook_path)
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"THERAPIST_DAILY recurring visual refresh failed: {exc}"
        ) from exc


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
        _refresh_daily_visuals(output)
        return report
    finally:
        prerequisite.unlink(missing_ok=True)
