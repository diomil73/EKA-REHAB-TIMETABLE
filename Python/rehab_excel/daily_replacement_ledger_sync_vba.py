from __future__ import annotations

MODULE_NAME = "modDailyReplacementLedgerSync"
MACRO_NAME = "SyncDailyReplacementLedgerFromFacts"

MODULE_CODE = r'''Option Explicit

Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const DAILY_SHEET As String = "THERAPIST_DAILY"

Public Function SyncDailyReplacementLedgerFromFacts() As Boolean
    Dim scriptPath As String, rowsPath As String, responsePath As String
    Dim commandLine As String, responseText As String
    Dim exitCode As Long

    On Error GoTo SyncError

    scriptPath = ResolveLedgerSyncTool()
    If Len(scriptPath) = 0 Then
        MsgBox "Δεν είναι διαθέσιμος ο συγχρονισμός της λίστας αλλαγών.", vbCritical, "Αλλαγές ημέρας"
        Exit Function
    End If

    ThisWorkbook.Save
    rowsPath = Environ$("TEMP") & "\\eka_replacement_ledger_rows_" & _
               Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\\eka_replacement_ledger_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --rows " & QuoteArg(rowsPath) & _
                  " --response " & QuoteArg(responsePath)

    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or _
       InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or _
       Dir$(rowsPath) = "" Then GoTo SyncFailed

    If Not ApplyLedgerRows(rowsPath) Then GoTo SyncFailed
    Application.Run "RefreshDailyChangeLedgerView"
    ThisWorkbook.Save
    SyncDailyReplacementLedgerFromFacts = True
    GoTo SyncExit

SyncFailed:
    WriteLedgerSyncDiagnostic exitCode, responseText
    MsgBox "Η αναπλήρωση αποθηκεύτηκε, αλλά η λίστα αλλαγών δεν συγχρονίστηκε σωστά." & vbCrLf & _
           "Μην κάνετε άλλη αλλαγή πριν ελεγχθεί το πρόβλημα.", vbCritical, "Αλλαγές ημέρας"
    GoTo SyncExit

SyncError:
    WriteLedgerSyncDiagnostic Err.Number, Err.Description
    MsgBox "Δεν ήταν δυνατός ο συγχρονισμός της λίστας αλλαγών.", vbCritical, "Αλλαγές ημέρας"

SyncExit:
    On Error Resume Next
    If Len(rowsPath) > 0 And Dir$(rowsPath) <> "" Then Kill rowsPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
End Function

Private Function ApplyLedgerRows(ByVal rowsPath As String) As Boolean
    Dim ws As Worksheet, wsDaily As Worksheet
    Dim rawText As String, lines As Variant, fields As Variant
    Dim rowIndex As Long, outputRow As Long, startRow As Long
    Dim previousEvents As Boolean, previousScreenUpdating As Boolean
    Dim dateText As String

    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LEDGER_SHEET
    End If

    If IsNumeric(ws.Range("R1").Value) Then startRow = CLng(ws.Range("R1").Value)
    If startRow < 2 Then startRow = Application.Max(22, wsDaily.UsedRange.Row + wsDaily.UsedRange.Rows.Count + 3)

    ws.Visible = xlSheetVisible
    ws.Unprotect
    ws.Cells.ClearContents
    WriteLedgerHeaders ws
    ws.Range("R1").Value2 = startRow

    rawText = ReadUtf8Text(rowsPath)
    lines = Split(Replace(rawText, vbCr, ""), vbLf)
    outputRow = 2
    For rowIndex = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(rowIndex)))) = 0 Then GoTo NextLine
        fields = Split(CStr(lines(rowIndex)), vbTab)
        If UBound(fields) <> 9 Then GoTo Failed
        ws.Cells(outputRow, 1).Value2 = PayloadValue(fields(0))
        dateText = PayloadValue(fields(1))
        If Len(dateText) <> 10 Then GoTo Failed
        ws.Cells(outputRow, 2).NumberFormat = "dd/mm/yyyy"
        ws.Cells(outputRow, 2).Value = DateSerial(CInt(Left$(dateText, 4)), CInt(Mid$(dateText, 6, 2)), CInt(Right$(dateText, 2)))
        ws.Cells(outputRow, 3).Value2 = PayloadValue(fields(2))
        ws.Cells(outputRow, 4).Value2 = PayloadValue(fields(3))
        ws.Cells(outputRow, 5).Value2 = PayloadValue(fields(4))
        ws.Cells(outputRow, 6).NumberFormat = "@"
        ws.Cells(outputRow, 6).Value2 = PayloadValue(fields(5))
        ws.Cells(outputRow, 7).Value2 = PayloadValue(fields(6))
        ws.Cells(outputRow, 8).NumberFormat = "@"
        ws.Cells(outputRow, 8).Value2 = PayloadValue(fields(7))
        ws.Cells(outputRow, 9).Value2 = PayloadValue(fields(8))
        ws.Cells(outputRow, 12).Value2 = UCase$(PayloadValue(fields(9)))
        outputRow = outputRow + 1
NextLine:
    Next rowIndex

    ws.Visible = xlSheetVeryHidden
    ws.Protect
    ApplyLedgerRows = True

ApplyExit:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Function
Failed:
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVeryHidden
        ws.Protect
    End If
    ApplyLedgerRows = False
    Resume ApplyExit
End Function

Private Sub WriteLedgerHeaders(ByVal ws As Worksheet)
    Dim headers As Variant, col As Long
    headers = Array( _
        "ChangeID", "Date", "PatientID", "Patient", "OriginalProvider", "OriginalTime", _
        "EffectiveProvider", "EffectiveTime", "Reason", "OriginalCell", "EffectiveCell", _
        "Status", "OriginalSnapshot", "EffectiveSnapshot", "ExpectedOriginal", "ExpectedEffective", _
        "RestoredAt", "Reserved", "DisplayRow")
    For col = 0 To UBound(headers)
        ws.Cells(1, col + 1).Value2 = CStr(headers(col))
    Next col
End Sub

Private Function ResolveLedgerSyncTool() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\\Python\\tools\\prepare_daily_replacement_ledger_sync.py"
    If Dir$(candidate) <> "" Then ResolveLedgerSyncTool = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\prepare_daily_replacement_ledger_sync.py"
    If Dir$(candidate) <> "" Then _
        ResolveLedgerSyncTool = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
End Function

Private Function PayloadValue(ByVal raw As String) As String
    Dim i As Long, ch As String, result As String, escaped As Boolean
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If escaped Then
            Select Case ch
                Case "t": result = result & vbTab
                Case "r": result = result & vbCr
                Case "n": result = result & vbLf
                Case "\\": result = result & "\\"
                Case Else: result = result & ch
            End Select
            escaped = False
        ElseIf ch = "\\" Then
            escaped = True
        Else
            result = result & ch
        End If
    Next i
    If escaped Then result = result & "\\"
    PayloadValue = result
End Function

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

Private Sub WriteLedgerSyncDiagnostic(ByVal errorNumber As Long, ByVal errorText As String)
    Dim filePath As String, handle As Integer
    filePath = Environ$("TEMP") & "\\eka_replacement_ledger_last_error.log"
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #handle, "error_number=" & CStr(errorNumber)
    Print #handle, "error=" & errorText
    Close #handle
End Sub
'''
MODULE_CODE = MODULE_CODE.replace("\\\\", "\\")


def install_daily_replacement_ledger_sync(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)
