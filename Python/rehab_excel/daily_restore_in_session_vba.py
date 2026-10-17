from __future__ import annotations


_RUN_START = "Private Function RunAuthoritativeRestore(ByVal replacementArgs As String, ByVal restoreAll As Boolean) As Boolean"
_JSON_READER_START = "Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String) As String"
_LEDGER_REASON_ASSIGNMENT = "wsDaily.Cells(displayRow, 5).Value = wsLedger.Cells(ledgerRow, 9).Value"
_LOCALIZED_LEDGER_REASON_ASSIGNMENT = "wsDaily.Cells(displayRow, 5).Value = GreekLedgerReason(CStr(wsLedger.Cells(ledgerRow, 9).Value))"


_IN_SESSION_BLOCK = r'''Private Function RunAuthoritativeRestore(ByVal replacementArgs As String, ByVal restoreAll As Boolean) As Boolean
    Dim scriptPath As String
    Dim outputPath As String
    Dim responsePath As String
    Dim commandLine As String
    Dim shell As Object
    Dim exitCode As Long
    Dim responseText As String
    Dim errorText As String
    Dim previewPath As String

    On Error GoTo RestoreError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Δεν μπορεί να γίνει επαναφορά.", vbExclamation, "Επαναφορά"
        Exit Function
    End If

    scriptPath = ResolveAuthoritativeRestoreScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Δεν βρέθηκε το authoritative εργαλείο επαναφοράς.", vbCritical, "Επαναφορά"
        Exit Function
    End If

    ThisWorkbook.Save

    outputPath = ThisWorkbook.Path & "\AUTHORITATIVE_LEDGER_RESTORE_PREVIEW.xlsm"
    responsePath = Environ$("TEMP") & "\eka_daily_restore_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    On Error Resume Next
    If Dir$(outputPath) <> "" Then Kill outputPath
    If Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo RestoreError

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --output " & QuoteArg(outputPath) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --overwrite" & replacementArgs
    If restoreAll Then commandLine = commandLine & " --all-active"

    Application.StatusBar = "Επαναφορά αλλαγής και επανυπολογισμός..."
    Set shell = CreateObject("WScript.Shell")
    exitCode = CLng(shell.Run(commandLine, 0, True))

    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)

    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        errorText = JsonStringValue(responseText, "error")
        If Len(errorText) = 0 Then errorText = "Άγνωστο σφάλμα authoritative επαναφοράς."
        MsgBox "Η επαναφορά δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & errorText, vbExclamation, "Επαναφορά"
        GoTo FailedCleanup
    End If

    previewPath = JsonStringValue(responseText, "output_path")
    If Len(Trim$(previewPath)) = 0 Or Dir$(previewPath) = "" Then
        MsgBox "Το backend δεν επέστρεψε έγκυρο verified preview.", vbCritical, "Επαναφορά"
        GoTo FailedCleanup
    End If

    ApplyAuthoritativeRestorePreview previewPath

    MsgBox "Η επαναφορά ολοκληρώθηκε επιτυχώς.", vbInformation, "Επαναφορά"

    RunAuthoritativeRestore = True

FailedCleanup:
    Application.StatusBar = False
    On Error Resume Next
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    If Len(outputPath) > 0 And Dir$(outputPath) <> "" Then Kill outputPath
    On Error GoTo 0
    Exit Function

RestoreError:
    Application.StatusBar = False
    MsgBox "Δεν ήταν δυνατή η authoritative επαναφορά: " & Err.Description, vbCritical, "Επαναφορά"
    Resume FailedCleanup
End Function

Private Sub ApplyAuthoritativeRestorePreview(ByVal previewPath As String)
    Dim previewWorkbook As Workbook
    Dim rollbackWorkbook As Workbook
    Dim rollbackPath As String
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousAlerts As Boolean
    Dim previousCalculation As XlCalculation
    Dim errorText As String

    On Error GoTo ApplyError

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousAlerts = Application.DisplayAlerts
    previousCalculation = Application.Calculation

    rollbackPath = Environ$("TEMP") & "\eka_daily_restore_rollback_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".xlsm"

    ThisWorkbook.SaveCopyAs rollbackPath

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set previewWorkbook = Application.Workbooks.Open( _
        Filename:=previewPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    HideWorkbookWindow previewWorkbook

    ValidateRestoreSheetSet previewWorkbook
    ValidateRestoreSheetSet ThisWorkbook
    SyncRestoreSheetSet previewWorkbook, ThisWorkbook
    ThisWorkbook.Save

    previewWorkbook.Close SaveChanges:=False
    Set previewWorkbook = Nothing

    On Error Resume Next
    If Dir$(rollbackPath) <> "" Then Kill rollbackPath
    On Error GoTo 0

ApplyExit:
    Application.Calculation = previousCalculation
    Application.DisplayAlerts = previousAlerts
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Sub

ApplyError:
    errorText = Err.Description
    On Error Resume Next
    If Not previewWorkbook Is Nothing Then previewWorkbook.Close SaveChanges:=False
    Set previewWorkbook = Nothing

    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then
        Set rollbackWorkbook = Application.Workbooks.Open( _
            Filename:=rollbackPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        HideWorkbookWindow rollbackWorkbook
        ValidateRestoreSheetSet rollbackWorkbook
        SyncRestoreSheetSet rollbackWorkbook, ThisWorkbook
        ThisWorkbook.Save
        rollbackWorkbook.Close SaveChanges:=False
        Set rollbackWorkbook = Nothing
        Kill rollbackPath
    End If
    On Error GoTo 0

    Application.Calculation = previousCalculation
    Application.DisplayAlerts = previousAlerts
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Err.Raise vbObjectError + 2910, "ApplyAuthoritativeRestorePreview", errorText
End Sub

Private Sub HideWorkbookWindow(ByVal workbookValue As Workbook)
    On Error Resume Next
    If workbookValue.Windows.Count > 0 Then workbookValue.Windows(1).Visible = False
    On Error GoTo 0
End Sub

Private Sub ValidateRestoreSheetSet(ByVal workbookValue As Workbook)
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet

    sheetNames = RestoreSheetNames()
    For Each sheetName In sheetNames
        Set ws = Nothing
        On Error Resume Next
        Set ws = workbookValue.Worksheets(CStr(sheetName))
        On Error GoTo 0
        If ws Is Nothing Then
            Err.Raise vbObjectError + 2911, , _
                "Λείπει απαιτούμενο φύλλο επαναφοράς: " & CStr(sheetName)
        End If
    Next sheetName
End Sub

Private Sub SyncRestoreSheetSet(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook)
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRange As Range

    sheetNames = RestoreSheetNames()
    For Each sheetName In sheetNames
        Set sourceSheet = sourceWorkbook.Worksheets(CStr(sheetName))
        Set targetSheet = targetWorkbook.Worksheets(CStr(sheetName))
        Set sourceRange = MeaningfulRestoreRange(sourceSheet)
        sourceRange.Copy Destination:=targetSheet.Range(sourceRange.Address)
    Next sheetName
    Application.CutCopyMode = False
End Sub

Private Function MeaningfulRestoreRange(ByVal ws As Worksheet) As Range
    Dim lastRowCell As Range
    Dim lastColumnCell As Range

    On Error Resume Next
    Set lastRowCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    Set lastColumnCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If lastRowCell Is Nothing Or lastColumnCell Is Nothing Then
        Set MeaningfulRestoreRange = ws.Range("A1")
    Else
        Set MeaningfulRestoreRange = ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(lastRowCell.Row, lastColumnCell.Column))
    End If
End Function

Private Function RestoreSheetNames() As Variant
    RestoreSheetNames = Array( _
        "DAILY_REPLACEMENT_FACTS", _
        "DAILY_CHANGE_LEDGER", _
        "THERAPIST_DAILY", _
        "MASTER_SCHEDULE", _
        "DAILY_STATISTICS")
End Function

Private Function ResolveAuthoritativeRestoreScript() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\Python\tools\restore_daily_replacement_facts.py"
    If Dir$(candidate) <> "" Then
        ResolveAuthoritativeRestoreScript = candidate
        Exit Function
    End If

    candidate = ThisWorkbook.Path & "\..\..\Python\tools\restore_daily_replacement_facts.py"
    If Dir$(candidate) <> "" Then
        ResolveAuthoritativeRestoreScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
    End If
End Function

Private Function GreekLedgerReason(ByVal rawReason As String) As String
    Dim normalized As String

    normalized = LCase$(Trim$(rawReason))
    Select Case normalized
        Case "", "temporary replacement", "daily_input therapist absence", "daily input therapist absence"
            GreekLedgerReason = "Απουσία θεραπευτή"
        Case Else
            GreekLedgerReason = rawReason
    End Select
End Function

'''


def build_in_session_ledger_vba(base_code: str) -> str:
    """Replace the obsolete close/reopen commit leg with an in-session bridge.

    The authoritative Python backend still builds and verifies a separate preview.
    Excel then imports the verified operational sheet state into the already-open
    workbook, with an in-session rollback copy. No terminal window, workbook close,
    or reopen worker is part of the primary ledger restore path.
    """

    start = base_code.find(_RUN_START)
    end = base_code.find(_JSON_READER_START)
    if start < 0 or end < 0 or end <= start:
        raise ValueError("Daily ledger VBA bridge markers were not found")
    if _LEDGER_REASON_ASSIGNMENT not in base_code:
        raise ValueError("Daily ledger reason assignment marker was not found")

    transformed = base_code[:start] + _IN_SESSION_BLOCK + base_code[end:]
    transformed = transformed.replace(
        _LEDGER_REASON_ASSIGNMENT,
        _LOCALIZED_LEDGER_REASON_ASSIGNMENT,
        1,
    )
    forbidden = (
        "Set exec = shell.Exec(commandLine)",
        "ThisWorkbook.Close SaveChanges:=True",
        "StartAuthoritativeRestoreCommit",
        "ResolveAuthoritativeCommitWorkerScript",
        "authoritative_commit_worker_cli.py",
    )
    leftovers = [marker for marker in forbidden if marker in transformed]
    if leftovers:
        raise ValueError(f"Obsolete close/reopen bridge markers remain: {leftovers!r}")
    return transformed