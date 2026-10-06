from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

from .patient_registration_form_vba import PATIENT_FORM_NAME

BRIDGE_MODULE_NAME = "modPatientRegistrationBridge"
BRIDGE_SCRIPT_RELATIVE = r"Python\tools\registration_transaction_worker_cli.py"

BRIDGE_MODULE_CODE = r'''Option Explicit
'''

FORM_BRIDGE_CODE = r'''
Private Sub cmdSave_Click()
    Dim requestPath As String
    Dim responsePath As String
    Dim workerScript As String
    Dim commandLine As String
    Dim patientType As String

    If Not ValidateForm() Then Exit Sub

    workerScript = ResolveRegistrationTransactionWorkerScript()
    If Len(workerScript) = 0 Then
        MsgBox "Δεν βρέθηκε ο worker ασφαλούς καταχώρησης του συστήματος.", _
               vbCritical, "Νέος ασθενής"
        Exit Sub
    End If

    On Error GoTo BridgeError
    ThisWorkbook.Save

    requestPath = Environ$("TEMP") & "\eka_registration_transaction_request_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 100)) & ".json"
    responsePath = Environ$("TEMP") & "\eka_registration_transaction_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 100)) & ".json"
    patientType = Trim$(cboPatientType.Value)

    WriteUtf8Text requestPath, BuildPatientRegistrationJson(patientType)

    MsgBox "Η καταχώρηση ξεκίνησε." & vbCrLf & vbCrLf & _
           "Το αρχείο θα κλείσει προσωρινά και θα ανοίξει ξανά αυτόματα " & _
           "μόλις ολοκληρωθεί η ασφαλής ενημέρωση.", _
           vbInformation, "Νέος ασθενής"

    ' The central registration menu is modal. If it remains loaded behind this
    ' form, Excel refuses an external Workbook.Close even after this form unloads.
    ' Release it before starting the detached transaction worker.
    On Error Resume Next
    Unload frmRegistrationMenu
    On Error GoTo BridgeError

    commandLine = QuoteArg("python") & " " & QuoteArg(workerScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    CreateObject("WScript.Shell").Run commandLine, 0, False
    Unload Me
    Exit Sub

BridgeError:
    MsgBox "Δεν ήταν δυνατή η εκκίνηση της ασφαλούς καταχώρησης: " & Err.Description, _
           vbCritical, "Νέος ασθενής"
End Sub

Private Function BuildPatientRegistrationJson(ByVal patientType As String) As String
    Dim roomValue As String
    Dim statusValue As String
    Dim infectiousValue As String
    Dim q As String

    q = Chr$(34)
    statusValue = Trim$(cboStatus.Value)

    If patientType = "Εξωτερικός" Then
        roomValue = ""
        infectiousValue = "false"
    Else
        roomValue = Trim$(cboRoom.Value)
        infectiousValue = LCase$(CStr(chkInfectious.Value))
    End If

    BuildPatientRegistrationJson = _
        "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_dir" & q & ":" & q & JsonEscape(ThisWorkbook.Path) & q & "," & _
        q & "action" & q & ":" & q & "new_patient" & q & "," & _
        q & "overwrite" & q & ":true," & _
        q & "values" & q & ":{" & _
        q & "patient_id" & q & ":" & q & q & "," & _
        q & "patient_type" & q & ":" & q & JsonEscape(patientType) & q & "," & _
        q & "hospital_mrn" & q & ":" & q & JsonEscape(Trim$(txtHospitalMRN.Text)) & q & "," & _
        q & "display_name" & q & ":" & q & JsonEscape(Trim$(txtDisplayName.Text)) & q & "," & _
        q & "room" & q & ":" & q & JsonEscape(roomValue) & q & "," & _
        q & "infectious" & q & ":" & infectiousValue & "," & _
        q & "status" & q & ":" & q & JsonEscape(statusValue) & q & _
        "}}"
End Function

Private Function ResolveRegistrationTransactionWorkerScript() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\Python\tools\registration_transaction_worker_cli.py"
    If Dir$(candidate) <> "" Then
        ResolveRegistrationTransactionWorkerScript = candidate
        Exit Function
    End If

    candidate = ThisWorkbook.Path & "\..\..\Python\tools\registration_transaction_worker_cli.py"
    If Dir$(candidate) <> "" Then
        ResolveRegistrationTransactionWorkerScript = _
            CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
    End If
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function JsonEscape(ByVal value As String) As String
    Dim text As String
    text = Replace(value, "\", "\\")
    text = Replace(text, Chr$(34), "\" & Chr$(34))
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    JsonEscape = text
End Function

Private Sub WriteUtf8Text(ByVal filePath As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
End Sub
'''


class PatientRegistrationBridgeVbaError(RuntimeError):
    """Raised when patient-form bridge wiring cannot be installed safely."""


@dataclass(frozen=True)
class PatientRegistrationBridgePreviewReport:
    source_path: str
    output_path: str
    source_unchanged: bool
    vba_present: bool
    form_present: bool
    bridge_wired: bool


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


def _replace_procedure(code_module, procedure_name: str, replacement: str) -> None:
    try:
        start_line = code_module.ProcStartLine(procedure_name, 0)
        line_count = code_module.ProcCountLines(procedure_name, 0)
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Procedure {procedure_name} was not found in {PATIENT_FORM_NAME}"
        ) from exc
    code_module.DeleteLines(start_line, line_count)
    code_module.InsertLines(start_line, replacement)


def _replace_bridge_helpers(code_module, helper_code: str) -> None:
    helper_names = (
        "ScheduleRegistrationClose",
        "BuildPatientRegistrationJson",
    )
    helper_start = 0
    for helper_name in helper_names:
        try:
            helper_start = int(code_module.ProcStartLine(helper_name, 0))
        except Exception:
            helper_start = 0
        if helper_start:
            break

    if helper_start:
        code_module.DeleteLines(helper_start, code_module.CountOfLines - helper_start + 1)

    code_module.AddFromString(helper_code)


def _replace_standard_module(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(BRIDGE_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = BRIDGE_MODULE_NAME
    module.CodeModule.AddFromString(BRIDGE_MODULE_CODE)


def wire_patient_registration_bridge(workbook_path: str | Path) -> None:
    if sys.platform != "win32":
        raise PatientRegistrationBridgeVbaError(
            "Patient registration bridge wiring requires Windows with Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientRegistrationBridgeVbaError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject
        component = vbproject.VBComponents(PATIENT_FORM_NAME)
        code_module = component.CodeModule

        bridge_text = FORM_BRIDGE_CODE
        save_end = bridge_text.index("Private Function BuildPatientRegistrationJson")
        save_proc = bridge_text[:save_end].rstrip()
        helper_code = bridge_text[save_end:].lstrip()

        _replace_procedure(code_module, "cmdSave_Click", save_proc)
        _replace_bridge_helpers(code_module, helper_code)
        _replace_standard_module(vbproject)

        stored_text = code_module.Lines(1, code_module.CountOfLines)
        required_markers = (
            "registration_transaction_worker_cli.py",
            "BuildPatientRegistrationJson",
            "ResolveRegistrationTransactionWorkerScript",
            'CreateObject("WScript.Shell").Run commandLine, 0, False',
        )
        if not all(marker in stored_text for marker in required_markers):
            raise PatientRegistrationBridgeVbaError(
                "Patient registration transaction wiring verification failed before workbook save"
            )

        workbook.Save()
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Patient registration bridge wiring failed: {exc}"
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
) -> PatientRegistrationBridgePreviewReport:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise PatientRegistrationBridgeVbaError(f"Source workbook not found: {source}")
    if source == output:
        raise PatientRegistrationBridgeVbaError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise PatientRegistrationBridgeVbaError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise PatientRegistrationBridgeVbaError(f"Output already exists: {output}")

    before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    try:
        wire_patient_registration_bridge(output)
    except Exception:
        output.unlink(missing_ok=True)
        raise

    after = _sha256(source)
    if before != after:
        output.unlink(missing_ok=True)
        raise PatientRegistrationBridgeVbaError("Source workbook changed during bridge wiring")

    if not _workbook_has_vba(output):
        output.unlink(missing_ok=True)
        raise PatientRegistrationBridgeVbaError("VBA project was not preserved")

    return PatientRegistrationBridgePreviewReport(
        source_path=str(source),
        output_path=str(output),
        source_unchanged=True,
        vba_present=True,
        form_present=True,
        bridge_wired=True,
    )
