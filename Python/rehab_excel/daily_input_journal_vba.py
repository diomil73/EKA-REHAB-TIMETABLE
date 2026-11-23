from __future__ import annotations


JOURNAL_MODULE_NAME = "modDailyInputJournal"

JOURNAL_MODULE_CODE = r'''Option Explicit

Private Const JOURNAL_SHEET As String = "DAILY_INPUT_JOURNAL"
Private Const JOURNAL_PREFIX As String = "diomil-journal-v1:"
Private Const XL_SHEET_VERY_HIDDEN As Long = 2

' The complete Apply/Restore state machine is deliberately not enabled yet.
' This gate prevents the legacy unjournaled Apply path from mutating the workbook.
Public Function DailyInputJournalApplyEnabled() As Boolean
    DailyInputJournalApplyEnabled = False
End Function

' One-shot technical probe. It writes only PREPARED and never calls Apply.
' Use only in a disposable synthetic workbook until all lifecycle phases exist.
Public Sub ProbeDailyInputPreparedJournal( _
    ByVal rowPath As String, _
    ByVal handoffPath As String, _
    ByVal payloadPath As String, _
    ByVal preflightScript As String, _
    ByVal readbackScript As String)

    If PersistPreparedDailyInputJournal( _
        rowPath, handoffPath, payloadPath, preflightScript, readbackScript) Then
        MsgBox "Η εγγραφή PREPARED αποθηκεύτηκε και επαληθεύτηκε." & vbCrLf & _
               "Δεν εφαρμόστηκαν αλλαγές σε ασθενείς ή θεραπείες.", _
               vbInformation, "Τεχνικός έλεγχος journal"
    Else
        MsgBox "Η εγγραφή PREPARED δεν επαληθεύτηκε." & vbCrLf & _
               "Δεν εφαρμόστηκαν αλλαγές. Μην συνεχίσετε καταχωρίσεις.", _
               vbCritical, "Τεχνικός έλεγχος journal"
    End If
End Sub

' VBA is the only writer of the operational XLSM. Python supplies a canonical
' five-field row; it never edits the workbook. The independent verifier checks
' the saved bytes and opens a second native Excel instance read-only.
Public Function PersistPreparedDailyInputJournal( _
    ByVal rowPath As String, _
    ByVal handoffPath As String, _
    ByVal payloadPath As String, _
    ByVal preflightScript As String, _
    ByVal readbackScript As String) As Boolean

    Dim ws As Worksheet, rowText As String, fields As Variant
    Dim lastRow As Long, targetRow As Long, candidateRow As Long
    Dim rowIndex As Long, columnIndex As Long
    Dim previousEvents As Boolean, saveAttempted As Boolean
    Dim commandLine As String, exitCode As Long
    Dim shell As Object

    On Error GoTo SafetyStop
    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    Application.StatusBar = "Επαλήθευση ασφαλούς εγγραφής journal..."

    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 2910, , "Workbook is read-only"
    If Len(Dir$(rowPath)) = 0 Or Len(Dir$(handoffPath)) = 0 Or _
       Len(Dir$(payloadPath)) = 0 Or Len(Dir$(preflightScript)) = 0 Or _
       Len(Dir$(readbackScript)) = 0 Then
        Err.Raise vbObjectError + 2911, , "Prepared evidence is missing"
    End If

    ' Reconfirm the exact saved source and complete prior journal prefix just
    ' before VBA appends. A stale workbook or changed payload cannot proceed.
    commandLine = QuoteJournalArg("python") & " " & QuoteJournalArg(preflightScript) & _
                  " --input " & QuoteJournalArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteJournalArg(payloadPath) & _
                  " --verify-handoff " & QuoteJournalArg(handoffPath)
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(commandLine, 0, True)
    If exitCode <> 0 Then Err.Raise vbObjectError + 2911, , "Prepared source recheck failed"

    rowText = Replace(ReadJournalUtf8(rowPath), vbCr, "")
    If Right$(rowText, 1) = vbLf Then rowText = Left$(rowText, Len(rowText) - 1)
    If Len(rowText) = 0 Or InStr(1, rowText, vbLf, vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 2912, , "Prepared row file must contain one line"
    End If
    fields = Split(rowText, vbTab)
    If UBound(fields) <> 4 Then Err.Raise vbObjectError + 2913, , "Prepared row shape invalid"
    If Len(Trim$(CStr(fields(0)))) = 0 Or CStr(fields(1)) <> "0" Or _
       CStr(fields(2)) <> "PREPARED" Then
        Err.Raise vbObjectError + 2914, , "Prepared row identity invalid"
    End If
    If Not IsJournalDigest(CStr(fields(3))) Then
        Err.Raise vbObjectError + 2915, , "Prepared row checksum invalid"
    End If
    If Left$(CStr(fields(4)), Len(JOURNAL_PREFIX)) <> JOURNAL_PREFIX Then
        Err.Raise vbObjectError + 2916, , "Prepared row payload prefix invalid"
    End If

    Set ws = FindJournalWorksheet()
    If ws Is Nothing Then Err.Raise vbObjectError + 2917, , "Existing journal unavailable"
    If int(ws.Visible) <> XL_SHEET_VERY_HIDDEN Or Not CBool(ws.ProtectContents) Then
        Err.Raise vbObjectError + 2918, , "Journal must be protected and veryHidden"
    End If
    For columnIndex = 1 To 5
        If CStr(ws.Cells(1, columnIndex).Value2) <> JournalHeader(columnIndex) Then
            Err.Raise vbObjectError + 2919, , "Journal header mismatch"
        End If
    Next columnIndex

    lastRow = 1
    For columnIndex = 1 To 5
        candidateRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
        If candidateRow > lastRow Then lastRow = candidateRow
    Next columnIndex
    For rowIndex = 2 To lastRow
        For columnIndex = 1 To 5
            If CBool(ws.Cells(rowIndex, columnIndex).HasFormula) Or _
               IsError(ws.Cells(rowIndex, columnIndex).Value2) Or _
               IsEmpty(ws.Cells(rowIndex, columnIndex).Value2) Then
                Err.Raise vbObjectError + 2920, , "Existing journal contains a partial row"
            End If
        Next columnIndex
        If CStr(ws.Cells(rowIndex, 1).Value2) = CStr(fields(0)) Then
            Err.Raise vbObjectError + 2921, , "Prepared operation ID already exists"
        End If
    Next rowIndex

    targetRow = lastRow + 1
    If targetRow > ws.Rows.Count Then Err.Raise vbObjectError + 2922, , "Journal row limit reached"
    For columnIndex = 1 To 5
        If Not IsEmpty(ws.Cells(targetRow, columnIndex).Value2) Then
            Err.Raise vbObjectError + 2923, , "Journal append target is not empty"
        End If
    Next columnIndex

    ws.Unprotect
    If CBool(ws.ProtectContents) Then Err.Raise vbObjectError + 2924, , "Journal unlock failed"
    ws.Cells(targetRow, 1).Value2 = CStr(fields(0))
    ws.Cells(targetRow, 2).Value2 = CLng(fields(1))
    ws.Cells(targetRow, 3).Value2 = CStr(fields(2))
    ws.Cells(targetRow, 4).Value2 = CStr(fields(3))
    ws.Cells(targetRow, 5).Value2 = CStr(fields(4))
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ws.Visible = XL_SHEET_VERY_HIDDEN
    If Not CBool(ws.ProtectContents) Or int(ws.Visible) <> XL_SHEET_VERY_HIDDEN Then
        Err.Raise vbObjectError + 2925, , "Journal protection readback failed"
    End If

    For columnIndex = 1 To 5
        If CStr(ws.Cells(targetRow, columnIndex).Value2) <> CStr(fields(columnIndex - 1)) Then
            Err.Raise vbObjectError + 2926, , "Prepared row in-memory readback mismatch"
        End If
    Next columnIndex

    saveAttempted = True
    ThisWorkbook.Save

    commandLine = QuoteJournalArg("python") & " " & QuoteJournalArg(readbackScript) & _
                  " --input " & QuoteJournalArg(ThisWorkbook.FullName) & _
                  " --handoff " & QuoteJournalArg(handoffPath) & _
                  " --payload " & QuoteJournalArg(payloadPath) & _
                  " --native-excel"
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(commandLine, 0, True)
    If exitCode <> 0 Then Err.Raise vbObjectError + 2927, , "Independent PREPARED readback failed"

    PersistPreparedDailyInputJournal = True
    GoTo CleanExit

SafetyStop:
    ' Never delete or rewrite a possibly durable journal row. Keep evidence and
    ' fail closed; a failed native readback must block every later operation.
    On Error Resume Next
    If Not ws Is Nothing Then
        If Not CBool(ws.ProtectContents) Then
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
        ws.Visible = XL_SHEET_VERY_HIDDEN
    End If
    PersistPreparedDailyInputJournal = False

CleanExit:
    Application.StatusBar = False
    Application.EnableEvents = previousEvents
End Function

Private Function FindJournalWorksheet() As Worksheet
    Dim index As Long, candidate As Worksheet
    For index = 1 To ThisWorkbook.Worksheets.Count
        Set candidate = ThisWorkbook.Worksheets(index)
        If StrComp(CStr(candidate.Name), JOURNAL_SHEET, vbBinaryCompare) = 0 Then
            Set FindJournalWorksheet = candidate
            Exit Function
        End If
    Next index
End Function

Private Function JournalHeader(ByVal columnIndex As Long) As String
    Select Case columnIndex
        Case 1: JournalHeader = "OperationID"
        Case 2: JournalHeader = "Revision"
        Case 3: JournalHeader = "Phase"
        Case 4: JournalHeader = "PayloadSHA256"
        Case 5: JournalHeader = "Payload"
    End Select
End Function

Private Function IsJournalDigest(ByVal value As String) As Boolean
    Dim index As Long, character As String
    If Len(value) <> 64 Then Exit Function
    For index = 1 To 64
        character = Mid$(value, index, 1)
        If InStr(1, "0123456789abcdef", character, vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsJournalDigest = True
End Function

Private Function ReadJournalUtf8(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadJournalUtf8 = stream.ReadText
    stream.Close
End Function

Private Function QuoteJournalArg(ByVal value As String) As String
    QuoteJournalArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def install_daily_input_journal(vbproject) -> None:
    """Install the isolated PREPARED bridge; Apply remains deliberately locked."""
    _remove_component_if_present(vbproject, JOURNAL_MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = JOURNAL_MODULE_NAME
    module.CodeModule.AddFromString(JOURNAL_MODULE_CODE)
