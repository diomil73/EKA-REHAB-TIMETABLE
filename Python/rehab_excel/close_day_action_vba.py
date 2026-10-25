from __future__ import annotations


CLOSE_DAY_MODULE_NAME = "modCloseDayAction"
CLOSE_DAY_MACRO_NAME = "CloseOperationalDay"


CLOSE_DAY_MODULE_CODE = r'''Option Explicit

Private mCloseDayHandoffPending As Boolean

Public Sub CloseOperationalDay()
    Dim scriptPath As String
    Dim outputPath As String
    Dim responsePath As String
    Dim commandLine As String
    Dim shell As Object
    Dim responseText As String
    Dim answer As VbMsgBoxResult
    Dim exitCode As Long

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Ανοίξτε το κανονικά και δοκιμάστε ξανά.", vbExclamation, "Κλείσιμο ημέρας"
        Exit Sub
    End If

    answer = MsgBox( _
        "Να κλείσει οριστικά η επιλεγμένη ημέρα;" & vbCrLf & vbCrLf & _
        "Μετά το κλείσιμο η ημέρα θα περάσει στο ιστορικό.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Κλείσιμο ημέρας")
    If answer <> vbYes Then Exit Sub

    scriptPath = ResolveCloseDayScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Η λειτουργία Κλεισίματος ημέρας δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
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
    exitCode = shell.Run(commandLine, 0, True)

    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η ημέρα δεν έκλεισε." & vbCrLf & vbCrLf & _
               "Ελέγξτε ότι δεν υπάρχουν εκκρεμείς συνεδρίες και δοκιμάστε ξανά." & vbCrLf & _
               "Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Κλείσιμο ημέρας"
        GoTo CleanUp
    End If

    If Not StartCloseDayCommit( _
        JsonStringValue(responseText, "output_path"), _
        JsonStringValue(responseText, "source_sha256_before")) Then
        GoTo CleanUp
    End If

    MsgBox "Η ημέρα έκλεισε επιτυχώς.", vbInformation, "Κλείσιμο ημέρας"

    On Error Resume Next
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0

    ScheduleDeferredCloseDayHandoff
    Exit Sub

CleanUp:
    On Error Resume Next
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0
    Exit Sub

ActionError:
    MsgBox "Η ημέρα δεν έκλεισε. Δοκιμάστε ξανά. Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
           vbCritical, "Κλείσιμο ημέρας"
End Sub

Private Sub ScheduleDeferredCloseDayHandoff()
    Dim callbackName As String

    If mCloseDayHandoffPending Then Exit Sub
    mCloseDayHandoffPending = True
    callbackName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!FinishCloseDayHandoff"

    On Error GoTo ScheduleError
    Application.OnTime EarliestTime:=Now + TimeSerial(0, 0, 1), Procedure:=callbackName, Schedule:=True
    Exit Sub

ScheduleError:
    mCloseDayHandoffPending = False
    MsgBox "Η ημέρα έκλεισε, αλλά το αρχείο δεν μπόρεσε να ολοκληρώσει την επανεκκίνηση. Κλείστε το Excel και ανοίξτε ξανά το αρχείο.", _
           vbExclamation, "Κλείσιμο ημέρας"
End Sub

Public Sub FinishCloseDayHandoff()
    On Error GoTo CloseError
    mCloseDayHandoffPending = False
    Application.EnableEvents = True
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub

CloseError:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "Η ημέρα έκλεισε, αλλά το αρχείο δεν μπόρεσε να κλείσει αυτόματα. Κλείστε το Excel και ανοίξτε ξανά το αρχείο.", _
           vbExclamation, "Κλείσιμο ημέρας"
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
        MsgBox "Δεν ήταν δυνατό να ολοκληρωθεί το Κλείσιμο ημέρας. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Κλείσιμο ημέρας"
        Exit Function
    End If
    If Len(Trim$(previewPath)) = 0 Or Len(Trim$(sourceSha256)) = 0 Then
        MsgBox "Δεν ήταν δυνατό να ολοκληρωθεί το Κλείσιμο ημέρας. Δοκιμάστε ξανά.", vbCritical, "Κλείσιμο ημέρας"
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
    MsgBox "Δεν ήταν δυνατό να ολοκληρωθεί το Κλείσιμο ημέρας. Επικοινωνήστε με τον διαχειριστή.", _
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
