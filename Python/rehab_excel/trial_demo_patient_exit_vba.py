"""Practice-DEMO-only patient discharge / outpatient programme end UI.

An explicit preview and confirmation records one date-effective exit in
PATIENT_EXITS. It NEVER deletes the registry/recurring schedule/history.
Only installed in an isolated fresh native DEMO and refuses any workbook whose
filename lacks _DEMO_. The independent Python source reader remains authoritative
for daily schedules; this UI does not claim completed native operator acceptance.
"""
from __future__ import annotations

from pathlib import Path
import sys

from .trial_demo_weekly_navigation import _shape


MODULE_NAME = "modDemoPatientExit"
BUTTON_NAME = "diomilDemoPatientExit"
MACRO_NAME = "DemoPatientExitAction"

VBA_CODE = r'''Option Explicit

Private Const EXIT_SHEET As String = "PATIENT_EXITS"

Private Function ExitLastRow(ByVal ws As Worksheet) As Long
    Dim col As Long, lastRow As Long, valueRow As Long
    lastRow = 1
    For col = 1 To 4
        valueRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If valueRow > lastRow Then lastRow = valueRow
    Next col
    ExitLastRow = lastRow
End Function

Private Function PatientRowForId(ByVal ws As Worksheet, ByVal wanted As String) As Long
    Dim rowNumber As Long, lastRow As Long, matches As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowNumber, 1).Value2)), wanted, vbTextCompare) = 0 Then
            matches = matches + 1
            PatientRowForId = rowNumber
        End If
    Next rowNumber
    If matches <> 1 Then PatientRowForId = 0
End Function

Private Function FindTypeColumn(ByVal ws As Worksheet) As Long
    Dim col As Long, labelText As String
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        labelText = Trim$(CStr(ws.Cells(1, col).Value2))
        If labelText = "PatientType" Or labelText = "ΤύποςΑσθενή" Or labelText = "Τύπος Ασθενή" Then
            FindTypeColumn = col
            Exit Function
        End If
    Next col
End Function

Private Function ParseExactIsoDate(ByVal value As String, ByRef result As Date) As Boolean
    Dim yr As Long, mo As Long, dy As Long
    On Error GoTo InvalidDate
    If Not value Like "####-##-##" Then Exit Function
    yr = CLng(Left$(value, 4))
    mo = CLng(Mid$(value, 6, 2))
    dy = CLng(Right$(value, 2))
    result = DateSerial(yr, mo, dy)
    If Year(result) <> yr Or Month(result) <> mo Or Day(result) <> dy Then Exit Function
    ParseExactIsoDate = True
InvalidDate:
End Function

Private Function ExpectedExitKind(ByVal label As String) As String
    If StrComp(label, "Εσωτερικός", vbTextCompare) = 0 Or _
       StrComp(label, "INPATIENT", vbTextCompare) = 0 Then
        ExpectedExitKind = "inpatient_discharge"
    ElseIf StrComp(label, "Εξωτερικός", vbTextCompare) = 0 Or _
           StrComp(label, "OUTPATIENT", vbTextCompare) = 0 Then
        ExpectedExitKind = "outpatient_program_end"
    End If
End Function

Private Function LedgerReady(ByVal exits As Worksheet, ByVal patients As Worksheet) As Boolean
    Dim col As Long, headers As Variant, rowNumber As Long
    Dim endRow As Long, pid As String, kind As String, lastValue As Variant
    Dim normalizedId As String, normalizedPrior As String, prior As Long
    Dim dateText As String, parsed As Date, patientRow As Long, typeCol As Long
    headers = Array("PatientID", "ExitKind", "LastActiveDate", "Reason")
    For col = 1 To 4
        If Trim$(CStr(exits.Cells(1, col).Value2)) <> CStr(headers(col - 1)) Then Exit Function
    Next col
    typeCol = FindTypeColumn(patients)
    If typeCol = 0 Then Exit Function
    endRow = ExitLastRow(exits)
    For rowNumber = 2 To endRow
        normalizedId = Trim$(CStr(exits.Cells(rowNumber, 1).Value2))
        kind = Trim$(CStr(exits.Cells(rowNumber, 2).Value2))
        If Len(normalizedId) = 0 Or Len(kind) = 0 Then Exit Function
        patientRow = PatientRowForId(patients, normalizedId)
        If patientRow = 0 Then Exit Function
        If kind <> ExpectedExitKind(CStr(patients.Cells(patientRow, typeCol).Value2)) Then Exit Function
        lastValue = exits.Cells(rowNumber, 3).Value
        If IsDate(lastValue) Then
            parsed = CDate(lastValue)
        Else
            dateText = Trim$(CStr(lastValue))
            If Not ParseExactIsoDate(dateText, parsed) Then Exit Function
        End If
        For prior = 2 To rowNumber - 1
            normalizedPrior = Trim$(CStr(exits.Cells(prior, 1).Value2))
            If StrComp(normalizedId, normalizedPrior, vbTextCompare) = 0 Then Exit Function
        Next prior
    Next rowNumber
    LedgerReady = True
End Function

Private Function RecurringCount(ByVal patientId As String) As Long
    Dim ws As Worksheet, rowNumber As Long, lastRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PROGRAM_WEEKLY_SLOTS")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For rowNumber = 2 To lastRow
            If StrComp(Trim$(CStr(ws.Cells(rowNumber, 1).Value2)), patientId, vbTextCompare) = 0 Then
                RecurringCount = RecurringCount + 1
            End If
        Next rowNumber
    End If
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For rowNumber = 2 To lastRow
            If StrComp(Trim$(CStr(ws.Cells(rowNumber, 1).Value2)), patientId, vbTextCompare) = 0 Then
                RecurringCount = RecurringCount + 1
            End If
        Next rowNumber
    End If
End Function

Public Sub DemoPatientExitAction()
    Dim patients As Worksheet, exits As Worksheet
    Dim patientId As String, dateInput As String, reason As String
    Dim patientRow As Long, typeCol As Long, exitRow As Long, rowNumber As Long
    Dim exitKind As String, patientName As String, serviceEnd As Date
    Dim oldEvents As Boolean, oldScreen As Boolean, written As Boolean
    Dim answer As VbMsgBoxResult, nBookings As Long
    Dim errText As String
    On Error GoTo Failed

    If InStr(1, ThisWorkbook.Name, "_DEMO_", vbTextCompare) = 0 Then
        MsgBox "Η ενέργεια επιτρέπεται μόνο σε ανεξάρτητο DEMO.", vbExclamation, "Diomil"
        Exit Sub
    End If
    If ThisWorkbook.ReadOnly Then
        MsgBox "Το DEMO έχει ανοίξει μόνο για ανάγνωση.", vbExclamation, "Diomil"
        Exit Sub
    End If

    Set patients = ThisWorkbook.Worksheets("PATIENTS")
    Set exits = ThisWorkbook.Worksheets(EXIT_SHEET)
    If exits.ProtectContents Then
        MsgBox "Το μητρώο εξιτηρίων είναι προστατευμένο.", vbExclamation, "Diomil"
        Exit Sub
    End If
    If Not LedgerReady(exits, patients) Then
        MsgBox "Οι εγγραφές εξιτηρίων ή οι ταυτότητες χρειάζονται έλεγχο.", vbCritical, "Diomil"
        Exit Sub
    End If
    typeCol = FindTypeColumn(patients)
    patientId = Trim$(InputBox("PatientID για εξιτήριο ή λήξη εξωτερικού προγράμματος:", "Diomil DEMO"))
    If Len(patientId) = 0 Then Exit Sub
    patientRow = PatientRowForId(patients, patientId)
    If patientRow = 0 Then
        MsgBox "Δεν βρέθηκε ακριβώς ένας ασθενής με αυτό το PatientID.", vbExclamation, "Diomil"
        Exit Sub
    End If
    exitKind = ExpectedExitKind(CStr(patients.Cells(patientRow, typeCol).Value2))
    If Len(exitKind) = 0 Then
        MsgBox "Ο τύπος ασθενούς δεν έχει επιβεβαιωθεί.", vbExclamation, "Diomil"
        Exit Sub
    End If
    exitRow = ExitLastRow(exits) + 1
    For rowNumber = 2 To exitRow - 1
        If StrComp(Trim$(CStr(exits.Cells(rowNumber, 1).Value2)), patientId, vbTextCompare) = 0 Then
            MsgBox "Υπάρχει ήδη καταχωρημένο εξιτήριο ή λήξη για αυτό το PatientID.", vbExclamation, "Diomil"
            Exit Sub
        End If
    Next rowNumber

    dateInput = Trim$(InputBox( _
        "Τελευταία ενεργή ημέρα (συμπεριλαμβάνεται), μορφή ΕΕΕΕ-ΜΜ-ΗΗ:", _
        "Diomil DEMO", Format$(Date, "yyyy-mm-dd")))
    If Len(dateInput) = 0 Then Exit Sub
    If Not ParseExactIsoDate(dateInput, serviceEnd) Then
        MsgBox "Η ημερομηνία πρέπει να είναι έγκυρη, π.χ. 2026-11-16.", vbExclamation, "Diomil"
        Exit Sub
    End If
    reason = Trim$(InputBox("Αιτιολογία (προαιρετική):", "Diomil DEMO"))
    patientName = Trim$(CStr(patients.Cells(patientRow, 3).Value2))
    nBookings = RecurringCount(patientId)
    answer = MsgBox( _
        "Ασθενής: " & patientName & vbCrLf & _
        "PatientID: " & patientId & vbCrLf & _
        "Ενέργεια: " & IIf(exitKind = "inpatient_discharge", "Οριστικό εξιτήριο", "Λήξη εξωτερικού προγράμματος") & vbCrLf & _
        "Τελευταία ενεργή ημέρα: " & dateInput & vbCrLf & _
        "Εβδομαδιαίες εγγραφές: " & CStr(nBookings) & vbCrLf & vbCrLf & _
        "Οι μεταγενέστερες συνεδρίες θα αποκλείονται από το ημερήσιο πρόγραμμα." & vbCrLf & _
        "Δεν θα διαγραφεί κανένα ιστορικό. Να αποθηκευτεί;", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Επιβεβαίωση εξιτηρίου")
    If answer <> vbYes Then Exit Sub

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Fail closed if a second action filled the target row since preview.
    For rowNumber = 1 To 4
        If Len(Trim$(CStr(exits.Cells(exitRow, rowNumber).Value2))) <> 0 Then
            Err.Raise vbObjectError + 2991, , "Ο επόμενος στίχος εξιτηρίου δεν είναι κενός."
        End If
    Next rowNumber
    exits.Cells(exitRow, 1).NumberFormat = "@"
    exits.Cells(exitRow, 1).Value = patientId
    exits.Cells(exitRow, 2).Value = exitKind
    exits.Cells(exitRow, 3).NumberFormat = "yyyy-mm-dd"
    exits.Cells(exitRow, 3).Value = serviceEnd
    exits.Cells(exitRow, 4).Value = reason
    written = True
    ThisWorkbook.Save
    If StrComp(Trim$(CStr(exits.Cells(exitRow, 1).Value2)), patientId, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 2992, , "Το PatientID δεν διατηρήθηκε μετά την αποθήκευση."
    End If
    If Not LedgerReady(exits, patients) Then
        Err.Raise vbObjectError + 2993, , "Ο επανέλεγχος εξιτηρίων απέτυχε."
    End If
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Το εξιτήριο/λήξη αποθηκεύτηκε." & vbCrLf & _
           "Το ημερήσιο πρόγραμμα θα εφαρμόζει την ημερομηνία λήξης." & vbCrLf & _
           "Το MASTER δεν ανανεώθηκε αυτόματα.", vbInformation, "Diomil DEMO"
    Exit Sub

Failed:
    errText = Err.Description
    On Error Resume Next
    If written Then
        exits.Range(exits.Cells(exitRow, 1), exits.Cells(exitRow, 4)).ClearContents
        ThisWorkbook.Save
    End If
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    MsgBox "Το εξιτήριο/λήξη ΔΕΝ επιβεβαιώθηκε. Ελέγξτε το αντίγραφο DEMO." & vbCrLf & _
           errText, vbCritical, "Diomil DEMO"
End Sub
'''


class DemoPatientExitVbaError(RuntimeError):
    """DEMO patient-exit UI was not installed or verified safely."""


def install_demo_patient_exit_action(workbook) -> None:
    """Install a separate DEMO-only VBA action without touching old buttons."""
    vbproject = workbook.VBProject
    try:
        component = vbproject.VBComponents(MODULE_NAME)
        count = int(component.CodeModule.CountOfLines)
        if count:
            component.CodeModule.DeleteLines(1, count)
    except Exception:
        component = vbproject.VBComponents.Add(1)
        component.Name = MODULE_NAME
    component.CodeModule.AddFromString(VBA_CODE)
    ws = workbook.Worksheets("MASTER_SCHEDULE")
    origin = ws.Range("A2")
    _shape(
        ws, name=BUTTON_NAME, caption="Εξιτήριο / Λήξη",
        action=MACRO_NAME,
        left=float(origin.Left) + 580, top=float(origin.Top) + 5,
        width=128,
    )


def verify_demo_patient_exit_action(path: str | Path) -> None:
    """Read-only COM verification, not an actual user/save/reopen acceptance."""
    if sys.platform != "win32":
        raise DemoPatientExitVbaError("Windows native Excel required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        raise DemoPatientExitVbaError("pywin32 required") from None
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(
            str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True,
        )
        module = wb.VBProject.VBComponents(MODULE_NAME)
        text = str(module.CodeModule.Lines(
            1, int(module.CodeModule.CountOfLines),
        ))
        if (
            "Public Sub DemoPatientExitAction()" not in text
            or "If Not LedgerReady(exits, patients) Then" not in text
            or "ThisWorkbook.Save" not in text
            or "vbDefaultButton2" not in text
        ):
            raise DemoPatientExitVbaError("Native exit macro contract missing")
        shape = wb.Worksheets("MASTER_SCHEDULE").Shapes(BUTTON_NAME)
        if str(shape.OnAction).split("!")[-1] != MACRO_NAME:
            raise DemoPatientExitVbaError("Native exit button action mismatch")
    except DemoPatientExitVbaError:
        raise
    except Exception:
        raise DemoPatientExitVbaError(
            "SAFETY STOP: native patient exit UI missing or unreadable"
        ) from None
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
