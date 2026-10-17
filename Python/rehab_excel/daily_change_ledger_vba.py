from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
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

    wsDaily.Columns("A").ColumnWidth = 4
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

    callerName = CStr(Application.Caller)
    rowIndex = CLng(Mid$(callerName, InStrRev(callerName, "_") + 1))
    If RestoreLedgerRow(rowIndex, True) Then RefreshDailyChangeLedgerView
End Sub

Public Sub RestoreSelectedDailyChanges()
    Dim wsDaily As Worksheet
    Dim shp As Shape
    Dim rowIndex As Long
    Dim restored As Long

    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    For Each shp In wsDaily.Shapes
        If Left$(shp.Name, Len(SHAPE_PREFIX & "Select_")) = SHAPE_PREFIX & "Select_" Then
            If shp.ControlFormat.Value = 1 Then
                rowIndex = CLng(Mid$(shp.Name, Len(SHAPE_PREFIX & "Select_") + 1))
                If RestoreLedgerRow(rowIndex, False) Then restored = restored + 1
            End If
        End If
    Next shp

    RefreshDailyChangeLedgerView
    If restored = 0 Then
        MsgBox "Δεν επιλέχθηκε ασφαλής αλλαγή για επαναφορά.", vbInformation
    Else
        MsgBox "Επαναφέρθηκαν " & CStr(restored) & " αλλαγές.", vbInformation
    End If
End Sub

Public Sub RestoreAllDailyChanges()
    Dim wsLedger As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim restored As Long

    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)
    lastRow = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row

    If MsgBox("Να επαναφερθούν όλες οι ενεργές αλλαγές της ημέρας;", vbYesNo + vbQuestion, "Επαναφορά όλων") <> vbYes Then Exit Sub

    For rowIndex = 2 To lastRow
        If Trim$(CStr(wsLedger.Cells(rowIndex, 12).Value)) = STATUS_ACTIVE Then
            If RestoreLedgerRow(rowIndex, False) Then restored = restored + 1
        End If
    Next rowIndex

    RefreshDailyChangeLedgerView
    MsgBox "Επαναφέρθηκαν " & CStr(restored) & " αλλαγές.", vbInformation
End Sub

Private Function RestoreLedgerRow(ByVal ledgerRow As Long, ByVal showMessage As Boolean) As Boolean
    Dim wsDaily As Worksheet
    Dim wsLedger As Worksheet
    Dim originalCell As String
    Dim effectiveCell As String
    Dim expectedOriginal As String
    Dim expectedEffective As String
    Dim currentOriginal As String
    Dim currentEffective As String

    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)
    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)

    If Trim$(CStr(wsLedger.Cells(ledgerRow, 12).Value)) <> STATUS_ACTIVE Then
        If showMessage Then MsgBox "Η αλλαγή έχει ήδη επαναφερθεί.", vbInformation
        Exit Function
    End If

    originalCell = Trim$(CStr(wsLedger.Cells(ledgerRow, 10).Value))
    effectiveCell = Trim$(CStr(wsLedger.Cells(ledgerRow, 11).Value))
    expectedOriginal = CStr(wsLedger.Cells(ledgerRow, 15).Value)
    expectedEffective = CStr(wsLedger.Cells(ledgerRow, 16).Value)
    currentOriginal = CStr(wsDaily.Range(originalCell).Value)
    currentEffective = CStr(wsDaily.Range(effectiveCell).Value)

    If currentOriginal <> expectedOriginal Then
        If showMessage Then MsgBox "Η αρχική θέση έχει αλλάξει ξανά. Δεν έγινε αυτόματη επαναφορά.", vbExclamation
        Exit Function
    End If
    If effectiveCell <> originalCell And currentEffective <> expectedEffective Then
        If showMessage Then MsgBox "Η νέα θέση έχει αλλάξει ξανά. Δεν έγινε αυτόματη επαναφορά.", vbExclamation
        Exit Function
    End If

    wsLedger.Cells(ledgerRow, 13).Copy Destination:=wsDaily.Range(originalCell)
    If effectiveCell <> originalCell Then
        wsLedger.Cells(ledgerRow, 14).Copy Destination:=wsDaily.Range(effectiveCell)
    End If

    wsLedger.Cells(ledgerRow, 12).Value = STATUS_RESTORED
    wsLedger.Cells(ledgerRow, 17).Value = Now
    RestoreLedgerRow = True

    If showMessage Then
        MsgBox "Η αλλαγή επαναφέρθηκε στην προηγούμενη θέση της.", vbInformation
    End If
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


def install_daily_change_ledger_preview(
    output_path: str | Path,
    source_path: str | Path,
    records: tuple[DailyChangeLedgerRecord, ...],
) -> None:
    """Install a native reversible ledger into a preview workbook.

    The source workbook is opened read-only only long enough to copy exact cell
    snapshots into a very-hidden ledger sheet. Restore buttons act only on the
    preview workbook and refuse to overwrite cells that have changed again.
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
                record.change_date,
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

        ledger.Visible = 2  # xlSheetVeryHidden

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
