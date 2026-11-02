from __future__ import annotations

DAY_AMENDMENT_MODULE_NAME = "modDayAmendmentAction"
DAY_AMENDMENT_MACRO_NAME = "StartClosedDayAmendment"
DAY_AMENDMENT_BUTTON_NAME = "ekaToolbar_DayAmendment"

DAY_AMENDMENT_MODULE_CODE = r'''Option Explicit

Public Sub StartClosedDayAmendment()
    Dim prepareScript As String, workerScript As String
    Dim payloadPath As String, prepareResponse As String, workerResponse As String
    Dim commandLine As String, responseText As String
    Dim actorOrSource As String
    Dim answer As VbMsgBoxResult, exitCode As Long

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Ανοίξτε το κανονικά και δοκιμάστε ξανά.", vbExclamation, "Αποθήκευση ημέρας"
        Exit Sub
    End If

    answer = MsgBox( _
        "Να αποθηκευτεί η κλεισμένη ημέρα με τις τρέχουσες αλλαγές;" & vbCrLf & vbCrLf & _
        "Η προηγούμενη καταχώρηση θα παραμείνει στο ιστορικό.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Αποθήκευση ημέρας")
    If answer <> vbYes Then Exit Sub

    prepareScript = ResolveRepoTool("prepare_day_amendment.py")
    workerScript = ResolveRepoTool("day_amendment_worker.py")
    If Len(prepareScript) = 0 Or Len(workerScript) = 0 Then
        MsgBox "Η λειτουργία Αποθήκευσης ημέρας δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Αποθήκευση ημέρας"
        Exit Sub
    End If

    ThisWorkbook.Save

    actorOrSource = "excel:" & Environ$("USERNAME")
    payloadPath = Environ$("TEMP") & "\\eka_day_amendment_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".json"
    prepareResponse = Environ$("TEMP") & "\\eka_day_amendment_prepare_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".json"
    workerResponse = Environ$("TEMP") & "\\eka_day_amendment_worker_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 100)) & ".json"

    ' Preflight while Excel is still open. This prevents closing the workbook if
    ' the selected day is not eligible for a new saved version.
    commandLine = QuoteArg("python") & " " & QuoteArg(prepareScript) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(prepareResponse) & _
                  " --actor-or-source " & QuoteArg(actorOrSource)

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(prepareResponse) <> "" Then responseText = ReadUtf8File(prepareResponse)

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or Dir$(payloadPath) = "" Then
        MsgBox "Η ημέρα δεν αποθηκεύτηκε." & vbCrLf & vbCrLf & _
               "Ελέγξτε ότι η ημέρα είναι κλεισμένη και ότι δεν υπάρχουν εκκρεμείς συνεδρίες." & vbCrLf & _
               "Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Αποθήκευση ημέρας"
        GoTo CleanUpPrepare
    End If

    ' Excel can rewrite the xlsm package during shutdown. The worker therefore
    ' rebuilds this payload after the file is fully released, preserving strict
    ' hash binding to the actual closed workbook before apply.
    commandLine = QuoteArg("python") & " " & QuoteArg(workerScript) & _
                  " --workbook " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(workerResponse) & _
                  " --refresh-payload-after-release" & _
                  " --actor-or-source " & QuoteArg(actorOrSource)

    CreateObject("WScript.Shell").Run commandLine, 0, False
    CloseForDaySaveHandoff
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
    MsgBox "Η ημέρα δεν αποθηκεύτηκε. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", vbCritical, "Αποθήκευση ημέρας"
End Sub

Private Sub CloseForDaySaveHandoff()
    Dim eventsWereEnabled As Boolean
    Dim workbookCount As Long

    On Error GoTo CloseError

    On Error Resume Next
    RestoreExcelInterface
    On Error GoTo CloseError

    eventsWereEnabled = Application.EnableEvents
    workbookCount = Application.Workbooks.Count
    Application.EnableEvents = False

    If workbookCount <= 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
        Application.EnableEvents = eventsWereEnabled
    End If
    Exit Sub

CloseError:
    On Error Resume Next
    Application.EnableEvents = eventsWereEnabled
    On Error GoTo 0
    Err.Raise vbObjectError + 731, "CloseForDaySaveHandoff", "Δεν ήταν δυνατό να κλείσει το αρχείο για την Αποθήκευση ημέρας."
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


def _last_used_row(ws) -> int:
    try:
        used = ws.UsedRange
        return max(1, int(used.Row) + int(used.Rows.Count) - 1)
    except Exception:
        return 1


def _delete_button_if_present(ws) -> None:
    try:
        ws.Shapes(DAY_AMENDMENT_BUTTON_NAME).Delete()
    except Exception:
        pass


def install_day_amendment_action(vbproject, workbook) -> None:
    try:
        component = vbproject.VBComponents(DAY_AMENDMENT_MODULE_NAME)
        vbproject.VBComponents.Remove(component)
    except Exception:
        pass

    module = vbproject.VBComponents.Add(1)
    module.Name = DAY_AMENDMENT_MODULE_NAME
    module.CodeModule.AddFromString(DAY_AMENDMENT_MODULE_CODE)

    # This installer owns only the day-save action on DAILY_INPUT.
    # MASTER and the wider application shell are intentionally untouched.
    ws = workbook.Worksheets("DAILY_INPUT")
    _delete_button_if_present(ws)

    last_row = _last_used_row(ws)
    anchor = ws.Range(f"A{last_row + 2}")
    left = float(anchor.Left) + 292
    top = float(anchor.Top) + 2
    shape = ws.Shapes.AddShape(5, left, top, 112, 24)
    shape.Name = DAY_AMENDMENT_BUTTON_NAME
    shape.OnAction = DAY_AMENDMENT_MACRO_NAME
    shape.AlternativeText = "Αποθήκευση αλλαγών σε ήδη κλεισμένη ημέρα με διατήρηση της προηγούμενης καταχώρησης στο ιστορικό."
    shape.Placement = 3
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.Line.Weight = 1
    shape.TextFrame2.TextRange.Text = "Αποθήκευση ημέρας"
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    try:
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    except Exception:
        pass
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
