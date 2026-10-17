from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path
import sys


LEDGER_SHEET = "DAILY_CHANGE_LEDGER"
MODULE_NAME = "modDailyChangeLedger"


@dataclass(frozen=True)
class DailyChangeLedgerRecord:
    change_id: str
    change_date: date
    patient_id: str
    patient_name: str
    original_provider: str
    original_time: time
    effective_provider: str
    effective_time: time
    original_cell: str
    effective_cell: str
    reason: str = "temporary replacement"

    @property
    def description(self) -> str:
        return (
            f"{self.patient_name}: {self.original_provider} "
            f"{self.original_time.strftime('%H:%M')} -> "
            f"{self.effective_provider} {self.effective_time.strftime('%H:%M')}"
        )


def _excel_date_value(value: date) -> datetime:
    """Return a pywin32/Excel-compatible midnight datetime for a Python date."""

    return datetime.combine(value, time.min)


VBA_CODE = r'''Option Explicit

Private Const LEDGER_SHEET As String = "DAILY_CHANGE_LEDGER"
Private Const DAILY_SHEET As String = "THERAPIST_DAILY"
Private Const SHAPE_PREFIX As String = "ekaLedger_"
Private Const STATUS_ACTIVE As String = "ACTIVE"
Private Const STATUS_RESTORED As String = "RESTORED"

Public Sub RefreshDailyChangeLedgerView()
    Dim wsDaily As Worksheet
    Dim wsLedger As Worksheet
    Dim startRow As Long
    Dim lastLedgerRow As Long
    Dim ledgerRow As Long
    Dim displayRow As Long
    Dim shp As Shape
    Dim cb As Shape
    Dim btn As Shape
    Dim titleText As String
    Dim statusText As String

    On Error GoTo RenderError

    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)
    startRow = CLng(wsLedger.Range("R1").Value)
    If startRow < 2 Then Exit Sub

    For Each shp In wsDaily.Shapes
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then shp.Delete
    Next shp

    wsDaily.Range("A" & startRow & ":H" & startRow + 100).ClearContents
    wsDaily.Range("A" & startRow & ":H" & startRow + 100).Interior.Pattern = xlNone
    wsDaily.Range("A" & startRow & ":H" & startRow + 100).Font.Bold = False
    wsDaily.Range("A" & startRow & ":H" & startRow + 100).Font.Italic = False

    lastLedgerRow = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row
    If lastLedgerRow < 2 Then Exit Sub

    titleText = "Αλλαγές για την " & Format$(wsLedger.Cells(2, 2).Value, "dd/mm/yyyy")
    With wsDaily.Range("A" & startRow & ":H" & startRow)
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
        .Font.Size = 12
    End With
    wsDaily.Cells(startRow, 1).Value = titleText

    wsDaily.Cells(startRow + 1, 1).Value = "✓"
    wsDaily.Cells(startRow + 1, 2).Value = "Ασθενής"
    wsDaily.Cells(startRow + 1, 3).Value = "Από"
    wsDaily.Cells(startRow + 1, 4).Value = "Προς"
    wsDaily.Cells(startRow + 1, 5).Value = "Αιτία"
    wsDaily.Cells(startRow + 1, 6).Value = "Κατάσταση"
    wsDaily.Cells(startRow + 1, 7).Value = "Ενέργεια"
    With wsDaily.Range("A" & startRow + 1 & ":G" & startRow + 1)
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    displayRow = startRow + 2
    For ledgerRow = 2 To lastLedgerRow
        If Len(Trim$(CStr(wsLedger.Cells(ledgerRow, 1).Value))) = 0 Then GoTo NextLedgerRow

        statusText = Trim$(CStr(wsLedger.Cells(ledgerRow, 12).Value))
        wsDaily.Cells(displayRow, 2).Value = wsLedger.Cells(ledgerRow, 4).Value
        wsDaily.Cells(displayRow, 3).Value = wsLedger.Cells(ledgerRow, 5).Value & " " & wsLedger.Cells(ledgerRow, 6).Text
        wsDaily.Cells(displayRow, 4).Value = wsLedger.Cells(ledgerRow, 7).Value & " " & wsLedger.Cells(ledgerRow, 8).Text
        wsDaily.Cells(displayRow, 5).Value = wsLedger.Cells(ledgerRow, 9).Value
        If statusText = STATUS_RESTORED Then
            wsDaily.Cells(displayRow, 6).Value = "Επαναφέρθηκε"
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.Color = RGB(128, 128, 128)
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.Italic = True
        Else
            wsDaily.Cells(displayRow, 6).Value = "Ενεργή αλλαγή"
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.ColorIndex = xlAutomatic
            wsDaily.Range("B" & displayRow & ":G" & displayRow).Font.Italic = False

            Set cb = wsDaily.Shapes.AddFormControl(1, _
                wsDaily.Cells(displayRow, 1).Left + 4, _
                wsDaily.Cells(displayRow, 1).Top + 2, 16, 16)
            cb.Name = SHAPE_PREFIX & "Select_" & CStr(ledgerRow)
            cb.TextFrame.Characters.Text = ""

            Set btn = wsDaily.Shapes.AddFormControl(0, _
                wsDaily.Cells(displayRow, 7).Left + 2, _
                wsDaily.Cells(displayRow, 7).Top + 1, _
                wsDaily.Cells(displayRow, 7).Width - 4, _
                wsDaily.Cells(displayRow, 7).Height - 2)
            btn.Name = SHAPE_PREFIX & "Restore_" & CStr(ledgerRow)
            btn.TextFrame.Characters.Text = "Επαναφορά"
            btn.OnAction = "RestoreDailyChangeRow"
        End If
        wsLedger.Cells(ledgerRow, 19).Value = displayRow
        displayRow = displayRow + 1
NextLedgerRow:
    Next ledgerRow

    Set btn = wsDaily.Shapes.AddFormControl(0, _
        wsDaily.Cells(displayRow + 1, 2).Left, _
        wsDaily.Cells(displayRow + 1, 2).Top, 170, 26)
    btn.Name = SHAPE_PREFIX & "RestoreSelected"
    btn.TextFrame.Characters.Text = "Επαναφορά επιλεγμένων"
    btn.OnAction = "RestoreSelectedDailyChanges"

    Set btn = wsDaily.Shapes.AddFormControl(0, _
        wsDaily.Cells(displayRow + 1, 5).Left, _
        wsDaily.Cells(displayRow + 1, 5).Top, 145, 26)
    btn.Name = SHAPE_PREFIX & "RestoreAll"
    btn.TextFrame.Characters.Text = "Επαναφορά όλων"
    btn.OnAction = "RestoreAllDailyChanges"

    wsDaily.Columns("A").ColumnWidth = 8
    wsDaily.Range("A2:A8,A12:A18").NumberFormat = "hh:mm"
    wsDaily.Range("A2:A8,A12:A18").HorizontalAlignment = xlCenter
    wsDaily.Columns("B").ColumnWidth = 28
    wsDaily.Columns("C:D").ColumnWidth = 23
    wsDaily.Columns("E").ColumnWidth = 28
    wsDaily.Columns("F").ColumnWidth = 18
    wsDaily.Columns("G").ColumnWidth = 20
    Exit Sub

RenderError:
    MsgBox "Δεν ήταν δυνατή η εμφάνιση της λίστας αλλαγών: " & Err.Description, vbExclamation
End Sub

Public Sub RestoreDailyChangeRow()
    Dim callerName As String
    Dim rowIndex As Long
    Dim replacementId As String
    Dim wsLedger As Worksheet

    On Error GoTo RestoreError
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)
    callerName = CStr(Application.Caller)
    rowIndex = CLng(Mid$(callerName, InStrRev(callerName, "_") + 1))

    If Trim$(CStr(wsLedger.Cells(rowIndex, 12).Value)) <> STATUS_ACTIVE Then
        MsgBox "Η αλλαγή έχει ήδη επαναφερθεί.", vbInformation, "Επαναφορά"
        Exit Sub
    End If

    replacementId = Trim$(CStr(wsLedger.Cells(rowIndex, 1).Value))
    If Len(replacementId) = 0 Then
        MsgBox "Δεν βρέθηκε το αναγνωριστικό της αλλαγής.", vbExclamation, "Επαναφορά"
        Exit Sub
    End If

    Call RunAuthoritativeRestore(" --replacement-id " & QuoteArg(replacementId), False)
    Exit Sub

RestoreError:
    MsgBox "Δεν ήταν δυνατή η έναρξη της επαναφοράς: " & Err.Description, vbCritical, "Επαναφορά"
End Sub

Public Sub RestoreSelectedDailyChanges()
    Dim wsDaily As Worksheet
    Dim wsLedger As Worksheet
    Dim shp As Shape
    Dim rowIndex As Long
    Dim replacementId As String
    Dim replacementArgs As String
    Dim selectedCount As Long

    On Error GoTo RestoreError
    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)

    For Each shp In wsDaily.Shapes
        If Left$(shp.Name, Len(SHAPE_PREFIX & "Select_")) = SHAPE_PREFIX & "Select_" Then
            If shp.ControlFormat.Value = 1 Then
                rowIndex = CLng(Mid$(shp.Name, Len(SHAPE_PREFIX & "Select_") + 1))
                If Trim$(CStr(wsLedger.Cells(rowIndex, 12).Value)) = STATUS_ACTIVE Then
                    replacementId = Trim$(CStr(wsLedger.Cells(rowIndex, 1).Value))
                    If Len(replacementId) > 0 Then
                        replacementArgs = replacementArgs & " --replacement-id " & QuoteArg(replacementId)
                        selectedCount = selectedCount + 1
                    End If
                End If
            End If
        End If
    Next shp

    If selectedCount = 0 Then
        MsgBox "Δεν επιλέχθηκε ενεργή αλλαγή για επαναφορά.", vbInformation, "Επαναφορά επιλεγμένων"
        Exit Sub
    End If

    Call RunAuthoritativeRestore(replacementArgs, False)
    Exit Sub

RestoreError:
    MsgBox "Δεν ήταν δυνατή η έναρξη της επαναφοράς: " & Err.Description, vbCritical, "Επαναφορά επιλεγμένων"
End Sub

Public Sub RestoreAllDailyChanges()
    If MsgBox("Να επαναφερθούν όλες οι ενεργές αλλαγές της ημέρας;", _
              vbYesNo + vbQuestion, "Επαναφορά όλων") <> vbYes Then Exit Sub

    Call RunAuthoritativeRestore("", True)
End Sub

Private Function RunAuthoritativeRestore(ByVal replacementArgs As String, ByVal restoreAll As Boolean) As Boolean
    Dim scriptPath As String
    Dim outputPath As String
    Dim responsePath As String
    Dim commandLine As String
    Dim shell As Object
    Dim exec As Object
    Dim stdoutText As String
    Dim stderrText As String
    Dim responseText As String
    Dim errorText As String

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
    On Error GoTo RestoreError

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --output " & QuoteArg(outputPath) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --overwrite" & replacementArgs
    If restoreAll Then commandLine = commandLine & " --all-active"

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(commandLine)
    Do While exec.Status = 0
        DoEvents
    Loop

    stdoutText = exec.StdOut.ReadAll
    stderrText = exec.StdErr.ReadAll
    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)

    If exec.ExitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        errorText = JsonStringValue(responseText, "error")
        If Len(errorText) = 0 Then errorText = Trim$(stderrText)
        If Len(errorText) = 0 Then errorText = Trim$(stdoutText)
        If Len(errorText) = 0 Then errorText = "Άγνωστο σφάλμα authoritative επαναφοράς."
        MsgBox "Η επαναφορά δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & errorText, vbExclamation, "Επαναφορά"
        GoTo FailedCleanup
    End If

    If Not StartAuthoritativeRestoreCommit( _
        JsonStringValue(responseText, "output_path"), _
        JsonStringValue(responseText, "source_sha256_before")) Then
        GoTo FailedCleanup
    End If

    MsgBox "Η επαναφορά επαληθεύτηκε και είναι έτοιμη για ασφαλή αποθήκευση." & vbCrLf & vbCrLf & _
           "Το αρχείο θα κλείσει προσωρινά και θα ανοίξει ξανά αυτόματα.", _
           vbInformation, "Επαναφορά"

    On Error Resume Next
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0

    RunAuthoritativeRestore = True
    ThisWorkbook.Close SaveChanges:=True
    Exit Function

FailedCleanup:
    On Error Resume Next
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    If Len(outputPath) > 0 And Dir$(outputPath) <> "" Then Kill outputPath
    On Error GoTo 0
    Exit Function

RestoreError:
    MsgBox "Δεν ήταν δυνατή η authoritative επαναφορά: " & Err.Description, vbCritical, "Επαναφορά"
    Resume FailedCleanup
End Function

Private Function StartAuthoritativeRestoreCommit(ByVal previewPath As String, ByVal sourceSha256 As String) As Boolean
    Dim workerScript As String
    Dim commitRequestPath As String
    Dim commitResponsePath As String
    Dim commandLine As String

    workerScript = ResolveAuthoritativeCommitWorkerScript()
    If Len(workerScript) = 0 Then
        MsgBox "Δεν βρέθηκε το authoritative save worker του συστήματος.", vbCritical, "Επαναφορά"
        Exit Function
    End If

    If Len(Trim$(previewPath)) = 0 Or Len(Trim$(sourceSha256)) = 0 Then
        MsgBox "Το backend δεν επέστρεψε τα στοιχεία ασφαλούς αποθήκευσης.", vbCritical, "Επαναφορά"
        Exit Function
    End If

    commitRequestPath = Environ$("TEMP") & "\eka_restore_commit_request_" & _
                        Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    commitResponsePath = Environ$("TEMP") & "\eka_restore_commit_response_" & _
                         Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    WriteUtf8Text commitRequestPath, BuildAuthoritativeRestoreCommitJson(previewPath, sourceSha256)

    commandLine = QuoteArg("python") & " " & QuoteArg(workerScript) & _
                  " --request " & QuoteArg(commitRequestPath) & _
                  " --response " & QuoteArg(commitResponsePath)

    On Error GoTo WorkerError
    CreateObject("WScript.Shell").Run commandLine, 0, False
    StartAuthoritativeRestoreCommit = True
    Exit Function

WorkerError:
    MsgBox "Δεν ήταν δυνατή η εκκίνηση της ασφαλούς αποθήκευσης: " & Err.Description, vbCritical, "Επαναφορά"
End Function

Private Function BuildAuthoritativeRestoreCommitJson(ByVal previewPath As String, ByVal sourceSha256 As String) As String
    Dim q As String
    q = Chr$(34)

    BuildAuthoritativeRestoreCommitJson = _
        "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_path" & q & ":" & q & JsonEscape(previewPath) & q & "," & _
        q & "expected_source_sha256" & q & ":" & q & JsonEscape(sourceSha256) & q & "," & _
        q & "remove_preview_after_success" & q & ":true," & _
        q & "reopen" & q & ":true," & _
        q & "timeout_seconds" & q & ":30," & _
        q & "poll_seconds" & q & ":0.5," & _
        q & "post_reopen_target" & q & ":" & q & JsonEscape(DAILY_SHEET) & q & _
        "}"
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

Private Function ResolveAuthoritativeCommitWorkerScript() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\Python\tools\authoritative_commit_worker_cli.py"
    If Dir$(candidate) <> "" Then
        ResolveAuthoritativeCommitWorkerScript = candidate
        Exit Function
    End If

    candidate = ThisWorkbook.Path & "\..\..\Python\tools\authoritative_commit_worker_cli.py"
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


class DailyChangeLedgerPreviewError(RuntimeError):
    pass


def _remove_component_if_present(vbproject, component_name: str) -> None:
    try:
        component = vbproject.VBComponents(component_name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _shape_caption(shape) -> str:
    """Return a shape caption when available without assuming one control type."""

    for reader in (
        lambda: shape.TextFrame2.TextRange.Text,
        lambda: shape.TextFrame.Characters().Text,
        lambda: shape.TextFrame.Characters.Text,
    ):
        try:
            value = reader()
        except Exception:
            continue
        if value is not None:
            return str(value).strip()
    return ""


def _remove_legacy_refresh_button(ws) -> int:
    """Remove only the obsolete THERAPIST_DAILY RefreshAll form button."""

    removed = 0
    try:
        buttons = ws.Buttons()
    except Exception:
        buttons = None

    if buttons is not None:
        for index in range(int(buttons.Count), 0, -1):
            button = buttons.Item(index)
            try:
                name = str(button.Name or "").strip().casefold()
            except Exception:
                name = ""
            try:
                caption = str(button.Caption or "").strip().casefold()
            except Exception:
                caption = ""
            try:
                action = str(button.OnAction or "").strip()
            except Exception:
                action = ""
            macro = action.rsplit("!", 1)[-1].strip().casefold()

            if (
                name == "button 1"
                and caption == "ανανέωση".casefold()
                and macro == "refreshall"
            ):
                button.Delete()
                removed += 1

    for index in range(int(ws.Shapes.Count), 0, -1):
        shape = ws.Shapes.Item(index)
        try:
            name = str(shape.Name or "").strip().casefold()
        except Exception:
            name = ""
        try:
            action = str(shape.OnAction or "").strip()
        except Exception:
            action = ""
        try:
            alt_text = str(shape.AlternativeText or "").strip().casefold()
        except Exception:
            alt_text = ""
        caption = _shape_caption(shape).strip().casefold() or alt_text
        macro = action.rsplit("!", 1)[-1].strip().casefold()
        if (
            name == "button 1"
            and caption == "ανανέωση".casefold()
            and macro == "refreshall"
        ):
            shape.Delete()
            removed += 1

    return removed


def install_daily_change_ledger_preview(
    output_path: str | Path,
    source_path: str | Path,
    records: tuple[DailyChangeLedgerRecord, ...],
) -> None:
    """Install a native reversible ledger into a preview workbook.

    The source workbook is opened read-only only long enough to copy exact cell
    snapshots into a very-hidden ledger sheet. Restore buttons now delegate to
    the authoritative replacement-fact backend; snapshots remain audit context.
    """

    output = Path(output_path).resolve()
    source = Path(source_path).resolve()
    if not records:
        raise DailyChangeLedgerPreviewError("At least one change record is required")
    if not output.exists() or not source.exists():
        raise DailyChangeLedgerPreviewError("Source and output workbooks must exist")
    if sys.platform != "win32":
        raise DailyChangeLedgerPreviewError("Windows Excel is required for ledger preview installation")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyChangeLedgerPreviewError("pywin32 is required") from exc

    excel = None
    source_wb = None
    output_wb = None
    failure: Exception | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        source_wb = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True)
        output_wb = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        source_daily = source_wb.Worksheets("THERAPIST_DAILY")
        output_daily = output_wb.Worksheets("THERAPIST_DAILY")
        _remove_legacy_refresh_button(output_daily)

        try:
            ledger = output_wb.Worksheets(LEDGER_SHEET)
            ledger.Cells.Clear()
        except Exception:
            ledger = output_wb.Worksheets.Add(After=output_wb.Worksheets(output_wb.Worksheets.Count))
            ledger.Name = LEDGER_SHEET

        headers = (
            "ChangeID", "Date", "PatientID", "Patient", "OriginalProvider", "OriginalTime",
            "EffectiveProvider", "EffectiveTime", "Reason", "OriginalCell", "EffectiveCell",
            "Status", "OriginalSnapshot", "EffectiveSnapshot", "ExpectedOriginal", "ExpectedEffective",
            "RestoredAt", "Reserved", "DisplayRow",
        )
        for col, header in enumerate(headers, start=1):
            ledger.Cells(1, col).Value = header

        last_used_row = int(output_daily.UsedRange.Row + output_daily.UsedRange.Rows.Count - 1)
        display_start_row = max(last_used_row + 3, 6)
        ledger.Range("R1").Value = display_start_row

        for row_index, record in enumerate(records, start=2):
            values = (
                record.change_id,
                _excel_date_value(record.change_date),
                record.patient_id,
                record.patient_name,
                record.original_provider,
                record.original_time.strftime("%H:%M"),
                record.effective_provider,
                record.effective_time.strftime("%H:%M"),
                record.reason,
                record.original_cell,
                record.effective_cell,
                "ACTIVE",
            )
            for col, value in enumerate(values, start=1):
                ledger.Cells(row_index, col).Value = value

            source_daily.Range(record.original_cell).Copy(Destination=ledger.Cells(row_index, 13))
            source_daily.Range(record.effective_cell).Copy(Destination=ledger.Cells(row_index, 14))
            ledger.Cells(row_index, 15).Value = output_daily.Range(record.original_cell).Value
            ledger.Cells(row_index, 16).Value = output_daily.Range(record.effective_cell).Value

        ledger.Visible = 2

        vbproject = output_wb.VBProject
        _remove_component_if_present(vbproject, MODULE_NAME)
        module = vbproject.VBComponents.Add(1)
        module.Name = MODULE_NAME
        module.CodeModule.AddFromString(VBA_CODE)

        output_wb.Save()
        excel.Run(f"'{output_wb.Name}'!RefreshDailyChangeLedgerView")
        output_wb.Save()
    except Exception as exc:
        failure = exc
    finally:
        if source_wb is not None:
            try:
                source_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if output_wb is not None:
            try:
                output_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    if failure is not None:
        raise DailyChangeLedgerPreviewError(str(failure)) from failure
