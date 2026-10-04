from __future__ import annotations

DAILY_INPUT_MODULE_NAME = "modDailyInputAction"
DAILY_INPUT_MACRO_NAME = "ApplyDailyInputPreview"

DAILY_INPUT_MODULE_CODE = r'''Option Explicit

Public Sub ApplyDailyInputPreview()
    Dim scriptPath As String
    Dim outputPath As String
    Dim commandLine As String
    Dim shell As Object
    Dim exec As Object
    Dim stdoutText As String
    Dim stderrText As String
    Dim responsePath As String
    Dim responseText As String

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Δεν μπορεί να αποθηκευτεί το DAILY_INPUT πριν από την εφαρμογή.", vbExclamation, "DAILY_INPUT"
        Exit Sub
    End If

    scriptPath = ResolveDailyInputScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Δεν βρέθηκε το εργαλείο εφαρμογής DAILY_INPUT.", vbCritical, "DAILY_INPUT"
        Exit Sub
    End If

    ThisWorkbook.Save

    outputPath = ThisWorkbook.Path & "\\DAILY_INPUT_APPLIED_PREVIEW.xlsm"
    responsePath = Environ$("TEMP") & "\\eka_daily_input_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --output " & QuoteArg(outputPath) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --overwrite"

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(commandLine)

    Do While exec.Status = 0
        DoEvents
    Loop

    stdoutText = exec.StdOut.ReadAll
    stderrText = exec.StdErr.ReadAll

    If Dir$(responsePath) <> "" Then
        responseText = ReadUtf8Text(responsePath)
    End If

    If exec.ExitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        If Len(JsonStringValue(responseText, "error")) > 0 Then
            MsgBox "Η εφαρμογή του DAILY_INPUT δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & _
                   JsonStringValue(responseText, "error"), vbExclamation, "DAILY_INPUT"
        ElseIf Len(Trim$(stderrText)) > 0 Then
            MsgBox "Η εφαρμογή του DAILY_INPUT δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & stderrText, vbExclamation, "DAILY_INPUT"
        Else
            MsgBox "Η εφαρμογή του DAILY_INPUT δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & stdoutText, vbExclamation, "DAILY_INPUT"
        End If
        GoTo CleanUp
    End If

    If JsonLongValue(responseText, "operational_absences") > 0 Then
        SetPostCommitTarget "REPLACEMENTS"
    Else
        ClearPostCommitTarget
    End If

    If Not StartAuthoritativeCommit( _
        JsonStringValue(responseText, "output_path"), _
        JsonStringValue(responseText, "source_sha256_before")) Then
        ClearPostCommitTarget
        GoTo CleanUp
    End If

    MsgBox "Η ημερήσια κατάσταση επαληθεύτηκε και είναι έτοιμη για αποθήκευση." & vbCrLf & vbCrLf & _
           "Το αρχείο θα κλείσει προσωρινά και θα ανοίξει ξανά αυτόματα μετά την ασφαλή αποθήκευση.", _
           vbInformation, "DAILY_INPUT"

    On Error Resume Next
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0

    ThisWorkbook.Close SaveChanges:=True
    Exit Sub

CleanUp:
    On Error Resume Next
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0
    Exit Sub

ActionError:
    MsgBox "Δεν ήταν δυνατή η εφαρμογή του DAILY_INPUT: " & Err.Description, vbCritical, "DAILY_INPUT"
End Sub

Private Function StartAuthoritativeCommit(ByVal previewPath As String, ByVal sourceSha256 As String) As Boolean
    Dim workerScript As String
    Dim commitRequestPath As String
    Dim commitResponsePath As String
    Dim commandLine As String

    workerScript = ResolveAuthoritativeCommitWorkerScript()
    If Len(workerScript) = 0 Then
        MsgBox "Δεν βρέθηκε το authoritative save worker του συστήματος.", vbCritical, "DAILY_INPUT"
        Exit Function
    End If

    If Len(Trim$(previewPath)) = 0 Or Len(Trim$(sourceSha256)) = 0 Then
        MsgBox "Το backend δεν επέστρεψε τα στοιχεία ασφαλούς αποθήκευσης.", vbCritical, "DAILY_INPUT"
        Exit Function
    End If

    commitRequestPath = Environ$("TEMP") & "\\eka_authoritative_commit_request_" & _
                        Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commitResponsePath = Environ$("TEMP") & "\\eka_authoritative_commit_response_" & _
                         Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    WriteUtf8Text commitRequestPath, BuildAuthoritativeCommitJson(previewPath, sourceSha256)

    commandLine = QuoteArg("python") & " " & QuoteArg(workerScript) & _
                  " --request " & QuoteArg(commitRequestPath) & _
                  " --response " & QuoteArg(commitResponsePath)

    On Error GoTo WorkerError
    CreateObject("WScript.Shell").Run commandLine, 0, False
    StartAuthoritativeCommit = True
    Exit Function

WorkerError:
    MsgBox "Δεν ήταν δυνατή η εκκίνηση της ασφαλούς αποθήκευσης: " & Err.Description, _
           vbCritical, "DAILY_INPUT"
End Function

Private Function BuildAuthoritativeCommitJson(ByVal previewPath As String, ByVal sourceSha256 As String) As String
    Dim q As String
    q = Chr$(34)

    BuildAuthoritativeCommitJson = _
        "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_path" & q & ":" & q & JsonEscape(previewPath) & q & "," & _
        q & "expected_source_sha256" & q & ":" & q & JsonEscape(sourceSha256) & q & "," & _
        q & "remove_preview_after_success" & q & ":true," & _
        q & "reopen" & q & ":true," & _
        q & "timeout_seconds" & q & ":30," & _
        q & "poll_seconds" & q & ":0.5" & _
        "}"
End Function

Private Function ResolveAuthoritativeCommitWorkerScript() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\\Python\\tools\\authoritative_commit_worker_cli.py"
    If Dir$(candidate) <> "" Then
        ResolveAuthoritativeCommitWorkerScript = candidate
        Exit Function
    End If

    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\authoritative_commit_worker_cli.py"
    If Dir$(candidate) <> "" Then
        ResolveAuthoritativeCommitWorkerScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
    End If
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

Private Function JsonLongValue(ByVal jsonText As String, ByVal key As String) As Long
    Dim marker As String, startPos As Long, endPos As Long, rawValue As String
    marker = Chr$(34) & key & Chr$(34) & ":"
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(marker)
    Do While startPos <= Len(jsonText) And Mid$(jsonText, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    endPos = startPos
    Do While endPos <= Len(jsonText) And Mid$(jsonText, endPos, 1) Like "[0-9]"
        endPos = endPos + 1
    Loop
    rawValue = Mid$(jsonText, startPos, endPos - startPos)
    If Len(rawValue) > 0 Then JsonLongValue = CLng(rawValue)
End Function

Private Sub SetPostCommitTarget(ByVal sheetName As String)
    On Error Resume Next
    ThisWorkbook.Names("__EKA_POST_COMMIT_TARGET").Delete
    ThisWorkbook.Names.Add Name:="__EKA_POST_COMMIT_TARGET", _
        RefersTo:="=" & Chr$(34) & sheetName & Chr$(34)
    On Error GoTo 0
End Sub

Private Sub ClearPostCommitTarget()
    On Error Resume Next
    ThisWorkbook.Names("__EKA_POST_COMMIT_TARGET").Delete
    On Error GoTo 0
End Sub

Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String) As String
    Dim marker As String, startPos As Long, index As Long, ch As String, escaped As Boolean, value As String
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

Private Function ResolveDailyInputScript() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\\Python\\tools\\apply_daily_input.py"
    If Dir$(candidate) <> "" Then
        ResolveDailyInputScript = candidate
        Exit Function
    End If

    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\apply_daily_input.py"
    If Dir$(candidate) <> "" Then
        ResolveDailyInputScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
    End If
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
'''


def install_daily_input_action(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(DAILY_INPUT_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = DAILY_INPUT_MODULE_NAME
    module.CodeModule.AddFromString(DAILY_INPUT_MODULE_CODE)
