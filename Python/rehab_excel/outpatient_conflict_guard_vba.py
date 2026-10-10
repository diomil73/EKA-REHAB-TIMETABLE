from __future__ import annotations

from .scheduling_conflict_popup_vba import install_scheduling_conflict_popup


FORM_NAME = "frmOutpatientSchedule"


OLD_SAVE_BLOCK = r'''Private Sub cmdSave_Click()
    Dim requestPath As String
    Dim responsePath As String
    Dim bridgeScript As String
    Dim commandLine As String
    Dim responseText As String
    Dim exitCode As Long

    If Not ValidateForm() Then Exit Sub

    bridgeScript = ResolveBridgeScript()
    If Len(bridgeScript) = 0 Then
        MsgBox "Δεν βρέθηκε το outpatient schedule bridge.", vbCritical
        Exit Sub
    End If

    requestPath = Environ$("TEMP") & "\eka_outpatient_schedule_request_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    responsePath = Environ$("TEMP") & "\eka_outpatient_schedule_response_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    WriteUtf8Text requestPath, BuildRequestJson()

    commandLine = QuoteArg("python") & " " & QuoteArg(bridgeScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo BridgeError
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) = "" Then
        MsgBox "Το backend δεν επέστρεψε αποτέλεσμα.", vbCritical
        GoTo CleanUp
    End If

    responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η καταχώρηση δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & JsonStringValue(responseText, "error"), vbExclamation
        GoTo CleanUp
    End If

    If Not StartAuthoritativeCommit( _
        JsonStringValue(responseText, "output_path"), _
        JsonStringValue(responseText, "source_sha256_before")) Then
        GoTo CleanUp
    End If

    MsgBox "Το πρόγραμμα επαληθεύτηκε και είναι έτοιμο για αποθήκευση." & vbCrLf & _
           "Base entry: " & JsonStringValue(responseText, "base_entry_id") & vbCrLf & vbCrLf & _
           "Το αρχείο θα κλείσει προσωρινά και θα ανοίξει ξανά αυτόματα μετά την ασφαλή αποθήκευση.", _
           vbInformation, "Πρόγραμμα εξωτερικού ασθενή"

    On Error Resume Next
    Kill requestPath
    Kill responsePath
    On Error GoTo 0

    Unload Me
    ThisWorkbook.Close SaveChanges:=True
    Exit Sub

CleanUp:
    On Error Resume Next
    Kill requestPath
    Kill responsePath
    On Error GoTo 0
    Exit Sub

BridgeError:
    MsgBox "Δεν ήταν δυνατή η εκτέλεση του backend: " & Err.Description, vbCritical
    Resume CleanUp
End Sub'''


NEW_SAVE_BLOCK = r'''Private Sub cmdSave_Click()
    Dim requestPath As String
    Dim responsePath As String
    Dim bridgeScript As String
    Dim commandLine As String
    Dim responseText As String
    Dim exitCode As Long
    Dim accepted As Boolean

    If Not ValidateForm() Then Exit Sub

    bridgeScript = ResolveBridgeScript()
    If Len(bridgeScript) = 0 Then
        MsgBox "Δεν βρέθηκε το outpatient schedule bridge.", vbCritical
        Exit Sub
    End If

    requestPath = Environ$("TEMP") & "\eka_outpatient_schedule_request_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    responsePath = Environ$("TEMP") & "\eka_outpatient_schedule_response_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(bridgeScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo BridgeError

    WriteUtf8Text requestPath, BuildRequestJson(False)
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) = "" Then
        MsgBox "Το backend δεν επέστρεψε αποτέλεσμα.", vbCritical
        GoTo CleanUp
    End If

    responseText = ReadUtf8Text(responsePath)
    If JsonStringValue(responseText, "conflict_type") = "therapist_double_booking" Then
        accepted = ConfirmTherapistDoubleBooking( _
            JsonStringValue(responseText, "therapist"), _
            JsonStringValue(responseText, "time"), _
            JsonStringValue(responseText, "existing_patient_name"), _
            JsonStringValue(responseText, "new_patient_name"))

        If Not accepted Then GoTo CleanUp

        WriteUtf8Text requestPath, BuildRequestJson(True)
        exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
        If Dir$(responsePath) = "" Then
            MsgBox "Το backend δεν επέστρεψε αποτέλεσμα μετά την επιβεβαίωση.", vbCritical
            GoTo CleanUp
        End If
        responseText = ReadUtf8Text(responsePath)
    End If

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η καταχώρηση δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & JsonStringValue(responseText, "error"), vbExclamation
        GoTo CleanUp
    End If

    If Not StartAuthoritativeCommit( _
        JsonStringValue(responseText, "output_path"), _
        JsonStringValue(responseText, "source_sha256_before")) Then
        GoTo CleanUp
    End If

    MsgBox "Το πρόγραμμα επαληθεύτηκε και είναι έτοιμο για αποθήκευση." & vbCrLf & _
           "Base entry: " & JsonStringValue(responseText, "base_entry_id") & vbCrLf & vbCrLf & _
           "Το αρχείο θα κλείσει προσωρινά και θα ανοίξει ξανά αυτόματα μετά την ασφαλή αποθήκευση.", _
           vbInformation, "Πρόγραμμα εξωτερικού ασθενή"

    On Error Resume Next
    Kill requestPath
    Kill responsePath
    On Error GoTo 0

    Unload Me
    ThisWorkbook.Close SaveChanges:=True
    Exit Sub

CleanUp:
    On Error Resume Next
    Kill requestPath
    Kill responsePath
    On Error GoTo 0
    Exit Sub

BridgeError:
    MsgBox "Δεν ήταν δυνατή η εκτέλεση του backend: " & Err.Description, vbCritical
    Resume CleanUp
End Sub'''


OLD_REQUEST_BLOCK = r'''Private Function BuildRequestJson() As String
    Dim q As String
    q = Chr$(34)
    BuildRequestJson = "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_dir" & q & ":" & q & JsonEscape(ThisWorkbook.Path) & q & "," & _
        q & "overwrite" & q & ":true," & _
        q & "values" & q & ":{" & _
        q & "patient_id" & q & ":" & q & JsonEscape(PatientIdFromSelection()) & q & "," & _
        q & "treatment" & q & ":" & q & JsonEscape(cboTreatment.Value) & q & "," & _
        q & "time" & q & ":" & q & JsonEscape(cboTime.Value) & q & "," & _
        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _
        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & _
        "}}"
End Function'''


NEW_REQUEST_BLOCK = r'''Private Function BuildRequestJson(ByVal allowDoubleBooking As Boolean) As String
    Dim q As String
    q = Chr$(34)
    BuildRequestJson = "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_dir" & q & ":" & q & JsonEscape(ThisWorkbook.Path) & q & "," & _
        q & "overwrite" & q & ":true," & _
        q & "values" & q & ":{" & _
        q & "patient_id" & q & ":" & q & JsonEscape(PatientIdFromSelection()) & q & "," & _
        q & "treatment" & q & ":" & q & JsonEscape(cboTreatment.Value) & q & "," & _
        q & "time" & q & ":" & q & JsonEscape(cboTime.Value) & q & "," & _
        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _
        q & "allow_double_booking" & q & ":" & LCase$(CStr(allowDoubleBooking)) & "," & _
        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & _
        "}}"
End Function'''


def patch_outpatient_schedule_form_code(code: str) -> str:
    if code.count(OLD_SAVE_BLOCK) != 1:
        raise ValueError("Expected exactly one outpatient save block to patch")
    if code.count(OLD_REQUEST_BLOCK) != 1:
        raise ValueError("Expected exactly one outpatient request builder to patch")
    return code.replace(OLD_SAVE_BLOCK, NEW_SAVE_BLOCK).replace(
        OLD_REQUEST_BLOCK, NEW_REQUEST_BLOCK
    )


def _position_control(control, left: float, top: float) -> None:
    control.Left = left
    control.Top = top


def install_outpatient_conflict_guard(vbproject) -> None:
    install_scheduling_conflict_popup(vbproject, position_control=_position_control)

    form = vbproject.VBComponents(FORM_NAME)
    module = form.CodeModule
    code = module.Lines(1, module.CountOfLines)
    patched = patch_outpatient_schedule_form_code(code)
    module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(patched)
