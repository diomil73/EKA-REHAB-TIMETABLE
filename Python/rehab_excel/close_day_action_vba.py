from __future__ import annotations


CLOSE_DAY_MODULE_NAME = "modCloseDayAction"
CLOSE_DAY_MACRO_NAME = "CloseOperationalDay"


CLOSE_DAY_MODULE_CODE = r'''Option Explicit

Public Sub CloseOperationalDay()
    Dim scriptPath As String
    Dim outputPath As String
    Dim responsePath As String
    Dim commandLine As String
    Dim shell As Object
    Dim exec As Object
    Dim stdoutText As String
    Dim stderrText As String
    Dim responseText As String
    Dim answer As VbMsgBoxResult

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Η ημέρα δεν μπορεί να κλείσει.", vbExclamation, "Κλείσιμο ημέρας"
        Exit Sub
    End If

    answer = MsgBox( _
        "Να κλείσει οριστικά η επιλεγμένη ημέρα;" & vbCrLf & vbCrLf & _
        "Το κλείσιμο θα αποθηκεύσει το τελικό ιστορικό και τα στατιστικά της ημέρας.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Κλείσιμο ημέρας")
    If answer <> vbYes Then Exit Sub

    scriptPath = ResolveCloseDayScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Δεν βρέθηκε το εργαλείο Κλεισίματος ημέρας.", vbCritical, "Κλείσιμο ημέρας"
        Exit Sub
    End If

    ThisWorkbook.Save

    outputPath = ThisWorkbook.Path & "\\CLOSE_DAY_VERIFIED_PREVIEW.xlsm"
    responsePath = Environ$("TEMP") & "\\eka_close_day_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --output " & QuoteArg(outputPath) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --actor-or-source " & QuoteArg("excel:" & Environ$("USERNAME")) & _
                  " --overwrite"

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(commandLine)

    Do While exec.Status = 0
        DoEvents
    Loop

    stdoutText = exec.StdOut.ReadAll
    stderrText = exec.StdErr.ReadAll
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)

    If exec.ExitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        If Len(JsonStringValue(responseText, "error")) > 0 Then
            MsgBox "Η ημέρα δεν έκλεισε:" & vbCrLf & vbCrLf & JsonStringValue(responseText, "error"), _
                   vbExclamation, "Κλείσιμο ημέρας"
        ElseIf Len(Trim$(stderrText)) > 0 Then
            MsgBox "Η ημέρα δεν έκλεισε:" & vbCrLf & vbCrLf & stderrText, vbExclamation, "Κλείσιμο ημέρας"
        Else
            MsgBox "Η ημέρα δεν έκλεισε:" & vbCrLf & vbCrLf & stdoutText, vbExclamation, "Κλείσιμο ημέρας"
        End If
        GoTo CleanUp
    End If

    If Not StartCloseDayCommit( _
        JsonStringValue(responseText, "output_path"), _
        JsonStringValue(responseText, "source_sha256_before")) Then
        GoTo CleanUp
    End If

    MsgBox "Η ημέρα επαληθεύτηκε και κλείστηκε επιτυχώς." & vbCrLf & vbCrLf & _
           "Η εφαρμογή θα ανοίξει ξανά αυτόματα στο MASTER μετά την ασφαλή αποθήκευση.", _
           vbInformation, "Κλείσιμο ημέρας"

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
    MsgBox "Δεν ήταν δυνατό το Κλείσιμο ημέρας: " & Err.Description, vbCritical, "Κλείσιμο ημέρας"
End Sub

Private Function StartCloseDayCommit(ByVal previewPath As String, ByVal sourceSha256 As String) As Boolean
    Dim workerScript As String
    Dim requestPath As String
    Dim responsePath As String
    Dim commandLine As String
    Dim q As String
    Dim payload As String

    workerScript = ResolveAuthoritativeCommitWorkerScript()
    If Len(workerScript) = 0 Then
        MsgBox "Δεν βρέθηκε το authoritative save worker του συστήματος.", vbCritical, "Κλείσιμο ημέρας"
        Exit Function
    End If
    If Len(Trim$(previewPath)) = 0 Or Len(Trim$(sourceSha256)) = 0 Then
        MsgBox "Το backend δεν επέστρεψε τα στοιχεία ασφαλούς αποθήκευσης.", vbCritical, "Κλείσιμο ημέρας"
        Exit Function
    End If

    requestPath = Environ$("TEMP") & "\\eka_close_day_commit_request_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    responsePath = Environ$("TEMP") & "\\eka_close_day_commit_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    q = Chr$(34)
    payload = "{" & _
              q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
              q & "preview_path" & q & ":" & q & JsonEscape(previewPath) & q & "," & _
              q & "expected_source_sha256" & q & ":" & q & JsonEscape(sourceSha256) & q & "," & _
              q & "remove_preview_after_success" & q & ":true," & _
              q & "reopen" & q & ":true," & _
              q & "timeout_seconds" & q & ":30," & _
              q & "poll_seconds" & q & ":0.5," & _
              q & "post_reopen_target" & q & ":" & q & "MASTER_SCHEDULE" & q & _
              "}"
    WriteUtf8Text requestPath, payload

    commandLine = QuoteArg("python") & " " & QuoteArg(workerScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)
    On Error GoTo WorkerError
    CreateObject("WScript.Shell").Run commandLine, 0, False
    StartCloseDayCommit = True
    Exit Function

WorkerError:
    MsgBox "Δεν ήταν δυνατή η εκκίνηση της ασφαλούς αποθήκευσης: " & Err.Description, _
           vbCritical, "Κλείσιμο ημέρας"
End Function

Private Function ResolveCloseDayScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\\Python\\tools\\close_day.py"
    If Dir$(candidate) <> "" Then
        ResolveCloseDayScript = candidate
        Exit Function
    End If
    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\close_day.py"
    If Dir$(candidate) <> "" Then
        ResolveCloseDayScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
    End If
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

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
'''


def install_close_day_action(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(CLOSE_DAY_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = CLOSE_DAY_MODULE_NAME
    module.CodeModule.AddFromString(CLOSE_DAY_MODULE_CODE)
