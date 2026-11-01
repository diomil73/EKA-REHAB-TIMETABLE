from __future__ import annotations

DAY_AMENDMENT_MODULE_NAME = "modDayAmendmentAction"
DAY_AMENDMENT_MACRO_NAME = "StartClosedDayAmendment"
DAY_AMENDMENT_BUTTON_NAME = "ekaToolbar_DayAmendment"

DAY_AMENDMENT_MODULE_CODE = r'''Option Explicit

Public Sub StartClosedDayAmendment()
    Dim prepareScript As String, workerScript As String
    Dim payloadPath As String, prepareResponse As String, workerResponse As String
    Dim commandLine As String, responseText As String
    Dim reasonCode As String, summaryText As String
    Dim answer As VbMsgBoxResult, exitCode As Long

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Ανοίξτε το κανονικά και δοκιμάστε ξανά.", vbExclamation, "Διόρθωση κλεισμένης ημέρας"
        Exit Sub
    End If

    answer = MsgBox( _
        "Να ξεκινήσει διόρθωση της κλεισμένης ημέρας;" & vbCrLf & vbCrLf & _
        "Η προηγούμενη έκδοση θα παραμείνει στο ιστορικό και θα δημιουργηθεί νέα έκδοση.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Διόρθωση κλεισμένης ημέρας")
    If answer <> vbYes Then Exit Sub

    reasonCode = Trim$(InputBox("Κωδικός / σύντομος λόγος διόρθωσης:", "Διόρθωση κλεισμένης ημέρας"))
    If Len(reasonCode) = 0 Then Exit Sub

    summaryText = Trim$(InputBox("Περιγραφή της διόρθωσης:", "Διόρθωση κλεισμένης ημέρας"))
    If Len(summaryText) = 0 Then Exit Sub

    prepareScript = ResolveRepoTool("prepare_day_amendment.py")
    workerScript = ResolveRepoTool("day_amendment_worker.py")
    If Len(prepareScript) = 0 Or Len(workerScript) = 0 Then
        MsgBox "Η λειτουργία διόρθωσης δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Διόρθωση κλεισμένης ημέρας"
        Exit Sub
    End If

    ThisWorkbook.Save

    payloadPath = Environ$("TEMP") & "\\eka_day_amendment_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".json"
    prepareResponse = Environ$("TEMP") & "\\eka_day_amendment_prepare_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".json"
    workerResponse = Environ$("TEMP") & "\\eka_day_amendment_worker_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(prepareScript) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(prepareResponse) & _
                  " --actor-or-source " & QuoteArg("excel:" & Environ$("USERNAME")) & _
                  " --reason-code " & QuoteArg(reasonCode) & _
                  " --summary " & QuoteArg(summaryText)

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(prepareResponse) <> "" Then responseText = ReadUtf8File(prepareResponse)

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then
        MsgBox "Η διόρθωση δεν προετοιμάστηκε." & vbCrLf & vbCrLf & _
               "Ελέγξτε ότι η ημέρα είναι κλεισμένη και ότι δεν υπάρχουν εκκρεμείς συνεδρίες." & vbCrLf & _
               "Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Διόρθωση κλεισμένης ημέρας"
        GoTo CleanUpPrepare
    End If

    commandLine = QuoteArg("python") & " " & QuoteArg(workerScript) & _
                  " --workbook " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(workerResponse)

    CreateObject("WScript.Shell").Run commandLine, 0, False

    Application.EnableEvents = False
    ThisWorkbook.Close SaveChanges:=False
    Application.EnableEvents = True
    Exit Sub

CleanUpPrepare:
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(prepareResponse) > 0 And Dir$(prepareResponse) <> "" Then Kill prepareResponse
    On Error GoTo 0
    Exit Sub

ActionError:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "Η διόρθωση δεν ολοκληρώθηκε. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", vbCritical, "Διόρθωση κλεισμένης ημέρας"
End Sub

Private Function ResolveRepoTool(ByVal fileName As String) As String
    Dim candidate As String, rootPath As String
    rootPath = ThisWorkbook.Path

    candidate = rootPath & "\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveRepoTool = candidate: Exit Function

    candidate = rootPath & "\\..\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveRepoTool = candidate: Exit Function

    candidate = rootPath & "\\..\\..\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then ResolveRepoTool = candidate: Exit Function
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function ReadUtf8File(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadUtf8File = stream.ReadText
    stream.Close
End Function
'''


def install_day_amendment_action(vbproject, workbook) -> None:
    try:
        component = vbproject.VBComponents(DAY_AMENDMENT_MODULE_NAME)
        vbproject.VBComponents.Remove(component)
    except Exception:
        pass

    module = vbproject.VBComponents.Add(1)
    module.Name = DAY_AMENDMENT_MODULE_NAME
    module.CodeModule.AddFromString(DAY_AMENDMENT_MODULE_CODE)

    ws = workbook.Worksheets("MASTER_SCHEDULE")
    try:
        ws.Shapes(DAY_AMENDMENT_BUTTON_NAME).Delete()
    except Exception:
        pass

    anchor = ws.Range("A3")
    shape = ws.Shapes.AddShape(5, float(anchor.Left) + 4, float(anchor.Top) + 2, 170, 24)
    shape.Name = DAY_AMENDMENT_BUTTON_NAME
    shape.OnAction = DAY_AMENDMENT_MACRO_NAME
    shape.AlternativeText = "Ρητή διόρθωση ήδη κλεισμένης ημέρας με διατήρηση όλων των προηγούμενων εκδόσεων."
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.Line.Weight = 1
    shape.TextFrame2.TextRange.Text = "Διόρθωση κλεισμένης ημέρας"
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    try:
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    except Exception:
        pass
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
