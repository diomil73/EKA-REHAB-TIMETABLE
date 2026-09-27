from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

from .patient_registration_form_vba import PATIENT_FORM_NAME

BRIDGE_MODULE_NAME = "modPatientRegistrationBridge"
BRIDGE_SCRIPT_RELATIVE = r"Python\tools\registration_bridge_cli.py"

FORM_BRIDGE_CODE = r'''
Private Sub cmdSave_Click()
    Dim requestPath As String
    Dim responsePath As String
    Dim previewDir As String
    Dim bridgeScript As String
    Dim pythonExe As String
    Dim commandLine As String
    Dim responseText As String
    Dim exitCode As Long
    Dim patientType As String

    If Not ValidateForm() Then Exit Sub

    bridgeScript = ResolveRegistrationBridgeScript()
    If Len(bridgeScript) = 0 Then
        MsgBox "Δεν βρέθηκε το registration bridge του συστήματος.", vbCritical, "Νέος ασθενής"
        Exit Sub
    End If

    pythonExe = "python"
    previewDir = ThisWorkbook.Path
    requestPath = Environ$("TEMP") & "\eka_registration_request_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    responsePath = Environ$("TEMP") & "\eka_registration_response_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    patientType = Trim$(cboPatientType.Value)

    WriteUtf8Text requestPath, BuildPatientRegistrationJson(patientType)

    commandLine = QuoteArg(pythonExe) & " " & QuoteArg(bridgeScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo BridgeError
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)

    If Dir$(responsePath) = "" Then
        MsgBox "Το registration backend δεν επέστρεψε αποτέλεσμα.", vbCritical, "Νέος ασθενής"
        GoTo CleanUp
    End If

    responseText = ReadUtf8Text(responsePath)

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η εγγραφή δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & _
               JsonStringValue(responseText, "error"), vbExclamation, "Νέος ασθενής"
        GoTo CleanUp
    End If

    txtPatientID.Text = JsonStringValue(responseText, "subject_key")
    MsgBox "Δημιουργήθηκε ασφαλές preview εγγραφής." & vbCrLf & _
           "Patient ID: " & txtPatientID.Text & vbCrLf & _
           "Αρχείο: " & JsonStringValue(responseText, "output_path"), _
           vbInformation, "Νέος ασθενής"

CleanUp:
    On Error Resume Next
    If Len(requestPath) > 0 Then Kill requestPath
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0
    Exit Sub

BridgeError:
    MsgBox "Δεν ήταν δυνατή η εκτέλεση του registration backend: " & Err.Description, _
           vbCritical, "Νέος ασθενής"
    Resume CleanUp
End Sub

Private Function BuildPatientRegistrationJson(ByVal patientType As String) As String
    Dim roomValue As String
    Dim statusValue As String
    Dim infectiousValue As String
    Dim q As String

    q = Chr$(34)

    If patientType = "Εξωτερικός" Then
        roomValue = ""
        statusValue = ""
        infectiousValue = "false"
    Else
        roomValue = Trim$(cboRoom.Value)
        statusValue = Trim$(cboStatus.Value)
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

Private Function ResolveRegistrationBridgeScript() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\Python\tools\registration_bridge_cli.py"
    If Dir$(candidate) <> "" Then
        ResolveRegistrationBridgeScript = candidate
        Exit Function
    End If

    candidate = ThisWorkbook.Path & "\..\..\Python\tools\registration_bridge_cli.py"
    If Dir$(candidate) <> "" Then
        ResolveRegistrationBridgeScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
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

Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String) As String
    Dim marker As String
    Dim startPos As Long
    Dim index As Long
    Dim ch As String
    Dim escaped As Boolean
    Dim value As String

    marker = Chr$(34) & key & Chr$(34) & ":"
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function

    startPos = InStr(startPos + Len(marker), jsonText, Chr$(34))
    If startPos = 0 Then Exit Function

    For index = startPos + 1 To Len(jsonText)
        ch = Mid$(jsonText, index, 1)
        If escaped Then
            Select Case ch
                Case "n": value = value & vbLf
                Case "r": value = value & vbCr
                Case "t": value = value & vbTab
                Case Else: value = value & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = Chr$(34) Then
            Exit For
        Else
            value = value & ch
        End If
    Next index

    JsonStringValue = value
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

Private Function ReadUtf8Text(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8Text = stream.ReadText
    stream.Close
End Function
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
    helper_name = "BuildPatientRegistrationJson"
    try:
        helper_start = code_module.ProcStartLine(helper_name, 0)
    except Exception:
        helper_start = 0

    if helper_start:
        code_module.DeleteLines(helper_start, code_module.CountOfLines - helper_start + 1)

    code_module.AddFromString(helper_code)


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

        stored_text = code_module.Lines(1, code_module.CountOfLines)
        required_markers = (
            "registration_bridge_cli.py",
            "BuildPatientRegistrationJson",
            "ResolveRegistrationBridgeScript",
            'Chr$(34) & "ok" & Chr$(34)',
        )
        if not all(marker in stored_text for marker in required_markers):
            raise PatientRegistrationBridgeVbaError(
                "Patient registration bridge verification failed before workbook save"
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
