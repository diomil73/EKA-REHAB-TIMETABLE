from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

from .patient_registration_form_vba import PATIENT_FORM_NAME

BRIDGE_MODULE_NAME = "modPatientRegistrationBridge"

BRIDGE_MODULE_CODE = r'''Option Explicit

Private Const OUTPATIENT_BLUE As Long = 16247773
Private Const INFECTIOUS_YELLOW As Long = 5101823

Public Function RegisterPatientInWorkbook( _
    ByVal patientType As String, _
    ByVal hospitalMRN As String, _
    ByVal displayName As String, _
    ByVal roomValue As String, _
    ByVal infectious As Boolean, _
    ByVal statusValue As String _
) As String
    Dim patients As Worksheet
    Dim planner As Worksheet
    Dim targetRow As Long
    Dim typeCol As Long
    Dim mrnCol As Long
    Dim doctorCol As Long
    Dim patientId As String
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo RegistrationError

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set patients = ThisWorkbook.Worksheets("PATIENTS")
    Set planner = ThisWorkbook.Worksheets("PATIENT_PLANNER")

    typeCol = FindOrCreateHeader(patients, Array("PatientType", "ΤύποςΑσθενή", "Τύπος Ασθενή"), "PatientType")
    mrnCol = FindOrCreateHeader(patients, Array("HospitalMRN", "ΑΜ Νοσοκομείου", "ΑΜΝοσοκομείου"), "HospitalMRN")
    doctorCol = FindOrCreateHeader(patients, Array("ResponsibleDoctor", "ΥπεύθυνοςΙατρός", "Υπεύθυνος Ιατρός", "ΥπεύθυνοςΓιατρός", "Υπεύθυνος Γιατρός"), "ResponsibleDoctor")

    ValidateFastRegistration patients, mrnCol, patientType, hospitalMRN, displayName, roomValue, statusValue

    patientId = NextPatientId(patients)
    targetRow = NextPatientTargetRow(patients)

    EnsurePatientTableRow patients, targetRow
    EnsureOutpatientScheduleSheet

    patients.Cells(targetRow, 1).Value = patientId
    If IsOutpatient(patientType) Then
        patients.Cells(targetRow, 2).Value = ""
        patients.Cells(targetRow, 4).Value = ""
    Else
        patients.Cells(targetRow, 2).Value = Trim$(roomValue)
        If infectious Then
            patients.Cells(targetRow, 4).Value = "Ν"
        Else
            patients.Cells(targetRow, 4).Value = ""
        End If
    End If
    patients.Cells(targetRow, 3).Value = Trim$(displayName)
    patients.Cells(targetRow, 5).Value = Trim$(statusValue)
    patients.Cells(targetRow, typeCol).Value = IIf(IsOutpatient(patientType), "Εξωτερικός", "Εσωτερικός")
    patients.Cells(targetRow, mrnCol).Value = Trim$(hospitalMRN)
    patients.Cells(targetRow, doctorCol).Value = ""

    RefreshPlannerRow planner, targetRow, IsOutpatient(patientType)

    If Not IsOutpatient(patientType) Then
        RebuildMasterFast patients, typeCol, doctorCol
    End If

    planner.Range("A" & CStr(targetRow) & ":E" & CStr(targetRow)).Calculate
    ThisWorkbook.Save

    RegisterPatientInWorkbook = patientId

RegistrationExit:
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    Exit Function

RegistrationError:
    Dim errorText As String
    errorText = Err.Description
    On Error Resume Next
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 2401, "RegisterPatientInWorkbook", errorText
End Function

Private Function IsOutpatient(ByVal patientType As String) As Boolean
    IsOutpatient = (LCase$(Trim$(patientType)) = LCase$("Εξωτερικός"))
End Function

Private Sub ValidateFastRegistration( _
    ByVal patients As Worksheet, _
    ByVal mrnCol As Long, _
    ByVal patientType As String, _
    ByVal hospitalMRN As String, _
    ByVal displayName As String, _
    ByVal roomValue As String, _
    ByVal statusValue As String _
)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim candidateMrn As String

    If Len(Trim$(patientType)) = 0 Then Err.Raise vbObjectError + 2410, , "Ο τύπος ασθενή είναι υποχρεωτικός."
    If Len(Trim$(displayName)) = 0 Then Err.Raise vbObjectError + 2411, , "Το ονοματεπώνυμο είναι υποχρεωτικό."
    If Len(Trim$(statusValue)) = 0 Then Err.Raise vbObjectError + 2412, , "Η κατάσταση είναι υποχρεωτική."
    If Not IsOutpatient(patientType) And Len(Trim$(roomValue)) = 0 Then Err.Raise vbObjectError + 2413, , "Ο θάλαμος είναι υποχρεωτικός για εσωτερικό ασθενή."

    candidateMrn = Trim$(hospitalMRN)
    If Len(candidateMrn) = 0 Then Exit Sub

    lastRow = LastNonBlankRow(patients, mrnCol)
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(patients.Cells(rowIndex, mrnCol).Value)), candidateMrn, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 2414, , "Ο ΑΜ Νοσοκομείου υπάρχει ήδη σε άλλον ασθενή."
        End If
    Next rowIndex
End Sub

Private Function FindOrCreateHeader(ByVal ws As Worksheet, ByVal candidates As Variant, ByVal canonicalName As String) As Long
    Dim lastCol As Long
    Dim col As Long
    Dim candidate As Variant
    Dim text As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5

    For col = 1 To lastCol
        text = Trim$(CStr(ws.Cells(1, col).Value))
        For Each candidate In candidates
            If StrComp(text, CStr(candidate), vbTextCompare) = 0 Then
                FindOrCreateHeader = col
                Exit Function
            End If
        Next candidate
    Next col

    FindOrCreateHeader = lastCol + 1
    ws.Cells(1, FindOrCreateHeader).Value = canonicalName
End Function

Private Function LastNonBlankRow(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    Dim found As Range
    On Error Resume Next
    Set found = ws.Columns(columnIndex).Find(What:="*", After:=ws.Cells(1, columnIndex), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If found Is Nothing Then
        LastNonBlankRow = 1
    Else
        LastNonBlankRow = found.Row
    End If
End Function

Private Function NextPatientTargetRow(ByVal ws As Worksheet) As Long
    Dim lastId As Long
    Dim lastName As Long
    lastId = LastNonBlankRow(ws, 1)
    lastName = LastNonBlankRow(ws, 3)
    NextPatientTargetRow = Application.Max(2, Application.Max(lastId, lastName) + 1)
End Function

Private Function NextPatientId(ByVal ws As Worksheet) As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rawValue As Variant
    Dim numericId As Long
    Dim highest As Long

    lastRow = LastNonBlankRow(ws, 1)
    highest = 0
    For rowIndex = 2 To lastRow
        rawValue = ws.Cells(rowIndex, 1).Value
        If Len(Trim$(CStr(rawValue))) > 0 And IsNumeric(rawValue) Then
            numericId = CLng(rawValue)
            If numericId > highest Then highest = numericId
        End If
    Next rowIndex
    NextPatientId = CStr(highest + 1)
End Function

Private Sub EnsurePatientTableRow(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim table As ListObject
    Dim firstDataRow As Long
    Dim lastDataRow As Long

    For Each table In ws.ListObjects
        If table.HeaderRowRange.Row = 1 And table.Range.Column <= 1 And (table.Range.Column + table.Range.Columns.Count - 1) >= 5 Then
            If table.DataBodyRange Is Nothing Then
                table.ListRows.Add
            End If
            firstDataRow = table.DataBodyRange.Row
            lastDataRow = firstDataRow + table.DataBodyRange.Rows.Count - 1
            If targetRow = lastDataRow + 1 Then table.ListRows.Add
            Exit For
        End If
    Next table
End Sub

Private Sub EnsureOutpatientScheduleSheet()
    Dim ws As Worksheet
    Dim headers As Variant
    Dim index As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "OUTPATIENT_SCHEDULE"
    End If

    headers = Array("PatientID", "Ασθενής", "Θεραπεία", "Ώρα", "Ημέρες", "Θεραπευτής")
    For index = LBound(headers) To UBound(headers)
        If Len(Trim$(CStr(ws.Cells(1, index + 1).Value))) = 0 Then
            ws.Cells(1, index + 1).Value = headers(index)
        ElseIf StrComp(Trim$(CStr(ws.Cells(1, index + 1).Value)), CStr(headers(index)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 2420, , "Το OUTPATIENT_SCHEDULE έχει μη συμβατή δομή στηλών."
        End If
    Next index
End Sub

Private Sub RefreshPlannerRow(ByVal planner As Worksheet, ByVal rowIndex As Long, ByVal outpatient As Boolean)
    planner.Cells(rowIndex, 1).Formula = "=IF(PATIENTS!C" & rowIndex & "=\"\",\"\",PATIENTS!A" & rowIndex & ")"
    planner.Cells(rowIndex, 2).Formula = "=IF(PATIENTS!C" & rowIndex & "=\"\",\"\",PATIENTS!B" & rowIndex & ")"
    planner.Cells(rowIndex, 3).Formula = "=IF(PATIENTS!C" & rowIndex & "=\"\",\"\",PATIENTS!C" & rowIndex & ")"
    planner.Cells(rowIndex, 4).Formula = "=IF(PATIENTS!C" & rowIndex & "=\"\",\"\",PATIENTS!D" & rowIndex & ")"
    planner.Cells(rowIndex, 5).Formula = "=IF(PATIENTS!C" & rowIndex & "=\"\",\"\",PATIENTS!E" & rowIndex & ")"

    If outpatient Then
        planner.Cells(rowIndex, 3).Interior.Color = RGB(221, 235, 247)
    Else
        planner.Cells(rowIndex, 3).Interior.Pattern = xlNone
    End If
End Sub

Private Function FindMasterHeaderRow(ByVal ws As Worksheet) As Long
    Dim rowIndex As Long
    For rowIndex = 1 To 10
        If StrComp(Trim$(CStr(ws.Cells(rowIndex, 2).Value)), "Θάλαμος", vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(ws.Cells(rowIndex, 3).Value)), "Ασθενής", vbTextCompare) = 0 Then
            FindMasterHeaderRow = rowIndex
            Exit Function
        End If
    Next rowIndex
    Err.Raise vbObjectError + 2430, , "Δεν βρέθηκε η κεφαλίδα του MASTER_SCHEDULE."
End Function

Private Function MasterStatusColumn(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    If StrComp(Trim$(CStr(ws.Cells(headerRow, 13).Value)), "Κατάσταση", vbTextCompare) = 0 Then
        MasterStatusColumn = 13
    ElseIf StrComp(Trim$(CStr(ws.Cells(headerRow, 11).Value)), "Κατάσταση", vbTextCompare) = 0 Then
        MasterStatusColumn = 11
    Else
        Err.Raise vbObjectError + 2431, , "Μη υποστηριζόμενη δομή MASTER_SCHEDULE."
    End If
End Function

Private Function ClinicPrefix(ByVal roomValue As String) As String
    Dim text As String
    text = LCase$(Trim$(roomValue))
    If Len(text) = 0 Then Exit Function
    If Left$(text, 1) = "a" Or Left$(text, 1) = "α" Then
        ClinicPrefix = "A"
    ElseIf Left$(text, 1) = "b" Or Left$(text, 1) = "β" Then
        ClinicPrefix = "B"
    End If
End Function

Private Function ClinicRank(ByVal roomValue As String) As Long
    Select Case ClinicPrefix(roomValue)
        Case "A": ClinicRank = 0
        Case "B": ClinicRank = 1
        Case Else: ClinicRank = 2
    End Select
End Function

Private Function RoomNumber(ByVal roomValue As String) As Long
    Dim index As Long
    Dim digits As String
    Dim ch As String
    For index = 1 To Len(roomValue)
        ch = Mid$(roomValue, index, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next index
    If Len(digits) = 0 Then
        RoomNumber = 999999
    Else
        RoomNumber = CLng(digits)
    End If
End Function

Private Function PatientRowComesAfter(ByVal patients As Worksheet, ByVal leftRow As Long, ByVal rightRow As Long) As Boolean
    Dim leftRoom As String
    Dim rightRoom As String
    Dim leftRank As Long
    Dim rightRank As Long
    Dim leftNumber As Long
    Dim rightNumber As Long
    Dim compareValue As Long

    leftRoom = Trim$(CStr(patients.Cells(leftRow, 2).Value))
    rightRoom = Trim$(CStr(patients.Cells(rightRow, 2).Value))
    leftRank = ClinicRank(leftRoom)
    rightRank = ClinicRank(rightRoom)

    If leftRank <> rightRank Then
        PatientRowComesAfter = (leftRank > rightRank)
        Exit Function
    End If

    leftNumber = RoomNumber(leftRoom)
    rightNumber = RoomNumber(rightRoom)
    If leftNumber <> rightNumber Then
        PatientRowComesAfter = (leftNumber > rightNumber)
        Exit Function
    End If

    compareValue = StrComp(Trim$(CStr(patients.Cells(leftRow, 3).Value)), Trim$(CStr(patients.Cells(rightRow, 3).Value)), vbTextCompare)
    If compareValue <> 0 Then
        PatientRowComesAfter = (compareValue > 0)
        Exit Function
    End If

    PatientRowComesAfter = (Val(patients.Cells(leftRow, 1).Value) > Val(patients.Cells(rightRow, 1).Value))
End Function

Private Sub SortPatientRows(ByVal patients As Worksheet, ByRef rows() As Long, ByVal count As Long)
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            If PatientRowComesAfter(patients, rows(i), rows(j)) Then
                temp = rows(i)
                rows(i) = rows(j)
                rows(j) = temp
            End If
        Next j
    Next i
End Sub

Private Function IsTruthy(ByVal value As Variant) As Boolean
    Dim text As String
    If VarType(value) = vbBoolean Then
        IsTruthy = CBool(value)
        Exit Function
    End If
    If IsNumeric(value) Then
        IsTruthy = (CDbl(value) <> 0)
        Exit Function
    End If
    text = LCase$(Trim$(CStr(value)))
    IsTruthy = (text = "ν" Or text = "ναι" Or text = "yes" Or text = "true" Or text = "1")
End Function

Private Sub ApplyMasterPatientStyle(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal infectious As Boolean)
    ws.Rows(rowIndex).RowHeight = 64
    ws.Range("B" & rowIndex & ":C" & rowIndex).Interior.Color = RGB(247, 247, 247)
    ws.Range("D" & rowIndex & ":D" & rowIndex).Interior.Color = RGB(217, 234, 247)
    ws.Range("E" & rowIndex & ":E" & rowIndex).Interior.Color = RGB(230, 228, 243)
    ws.Range("F" & rowIndex & ":F" & rowIndex).Interior.Color = RGB(216, 240, 244)
    ws.Range("G" & rowIndex & ":G" & rowIndex).Interior.Color = RGB(232, 239, 246)
    ws.Range("H" & rowIndex & ":H" & rowIndex).Interior.Color = RGB(228, 240, 220)
    ws.Range("I" & rowIndex & ":I" & rowIndex).Interior.Color = RGB(252, 240, 207)
    ws.Range("J" & rowIndex & ":J" & rowIndex).Interior.Color = RGB(237, 244, 227)
    ws.Range("K" & rowIndex & ":K" & rowIndex).Interior.Color = RGB(250, 229, 219)
    ws.Range("L" & rowIndex & ":L" & rowIndex).Interior.Color = RGB(240, 234, 247)
    ws.Range("M" & rowIndex & ":M" & rowIndex).Interior.Color = RGB(247, 247, 247)

    ws.Range("B" & rowIndex & ":M" & rowIndex).Font.Size = 10.5
    ws.Range("B" & rowIndex & ":M" & rowIndex).WrapText = True
    ws.Range("B" & rowIndex & ":M" & rowIndex).ShrinkToFit = True
    ws.Range("B" & rowIndex & ":M" & rowIndex).VerticalAlignment = xlCenter
    ws.Range("B" & rowIndex & ":M" & rowIndex).HorizontalAlignment = xlCenter
    ws.Cells(rowIndex, 2).Font.Bold = True
    ws.Cells(rowIndex, 3).Font.Bold = True
    ws.Cells(rowIndex, 3).HorizontalAlignment = xlLeft

    If infectious Then ws.Range("B" & rowIndex & ":C" & rowIndex).Interior.Color = RGB(255, 216, 77)
End Sub

Private Sub ApplyClinicBanner(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal clinic As String)
    Dim label As String
    If clinic = "A" Then
        label = "Α' ΚΛΙΝΙΚΗ"
    Else
        label = "Β' ΚΛΙΝΙΚΗ"
    End If

    With ws.Range("B" & rowIndex & ":M" & rowIndex)
        .ClearContents
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(rowIndex, 2).Value = label
    ws.Rows(rowIndex).RowHeight = 30
End Sub

Private Sub RebuildMasterFast(ByVal patients As Worksheet, ByVal typeCol As Long, ByVal doctorCol As Long)
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim statusCol As Long
    Dim lastPatientRow As Long
    Dim patientRows() As Long
    Dim patientCount As Long
    Dim rowIndex As Long
    Dim patientType As String
    Dim roomValue As String
    Dim targetRow As Long
    Dim oldLast As Long
    Dim clearLast As Long
    Dim clinic As String
    Dim previousClinic As String
    Dim plannerRow As Long
    Dim doctor As String
    Dim displayName As String

    Set ws = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    headerRow = FindMasterHeaderRow(ws)
    statusCol = MasterStatusColumn(ws, headerRow)
    lastPatientRow = Application.Max(LastNonBlankRow(patients, 1), LastNonBlankRow(patients, 3))

    ReDim patientRows(1 To Application.Max(1, lastPatientRow - 1))
    For rowIndex = 2 To lastPatientRow
        displayName = Trim$(CStr(patients.Cells(rowIndex, 3).Value))
        If Len(displayName) > 0 Then
            patientType = LCase$(Trim$(CStr(patients.Cells(rowIndex, typeCol).Value)))
            roomValue = Trim$(CStr(patients.Cells(rowIndex, 2).Value))
            If patientType = LCase$("Εσωτερικός") Or (Len(patientType) = 0 And Len(roomValue) > 0) Then
                patientCount = patientCount + 1
                patientRows(patientCount) = rowIndex
            End If
        End If
    Next rowIndex

    If patientCount > 1 Then SortPatientRows patients, patientRows, patientCount

    oldLast = Application.Max(LastNonBlankRow(ws, 2), LastNonBlankRow(ws, 3), LastNonBlankRow(ws, statusCol))
    clearLast = Application.Max(oldLast, headerRow + patientCount + 2)

    For rowIndex = headerRow + 1 To clearLast
        ws.Range("A" & rowIndex & ":P" & rowIndex).ClearContents
    Next rowIndex

    targetRow = headerRow + 1
    previousClinic = ""

    For rowIndex = 1 To patientCount
        plannerRow = patientRows(rowIndex)
        roomValue = Trim$(CStr(patients.Cells(plannerRow, 2).Value))
        clinic = ClinicPrefix(roomValue)

        If Len(clinic) > 0 And clinic <> previousClinic Then
            ApplyClinicBanner ws, targetRow, clinic
            targetRow = targetRow + 1
            previousClinic = clinic
        End If

        ws.Cells(targetRow, 1).Formula = "=PATIENT_PLANNER!D" & plannerRow
        ws.Cells(targetRow, 2).Formula = "=PATIENT_PLANNER!B" & plannerRow
        ws.Cells(targetRow, 4).Formula = "=PATIENT_PLANNER!F" & plannerRow
        ws.Cells(targetRow, 5).Formula = "=PATIENT_PLANNER!I" & plannerRow
        ws.Cells(targetRow, 6).Formula = "=PATIENT_PLANNER!N" & plannerRow
        ws.Cells(targetRow, 7).Formula = "=PATIENT_PLANNER!L" & plannerRow
        ws.Cells(targetRow, 8).Formula = "=PATIENT_PLANNER!P" & plannerRow
        ws.Cells(targetRow, 9).Formula = "=PATIENT_PLANNER!R" & plannerRow
        ws.Cells(targetRow, 10).Formula = "=PATIENT_PLANNER!T" & plannerRow
        ws.Cells(targetRow, statusCol).Formula = "=PATIENT_PLANNER!E" & plannerRow

        doctor = Trim$(CStr(patients.Cells(plannerRow, doctorCol).Value))
        displayName = Trim$(CStr(patients.Cells(plannerRow, 3).Value))
        If Len(doctor) > 0 Then
            ws.Cells(targetRow, 3).Value = displayName & vbLf & "✚ " & doctor
        Else
            ws.Cells(targetRow, 3).Formula = "=PATIENT_PLANNER!C" & plannerRow
        End If

        ApplyMasterPatientStyle ws, targetRow, IsTruthy(patients.Cells(plannerRow, 4).Value)
        targetRow = targetRow + 1
    Next rowIndex

    ws.Columns("A").Hidden = True
    On Error Resume Next
    ws.ScrollArea = "B1:M" & CStr(Application.Max(headerRow + 1, targetRow))
    On Error GoTo 0
    ws.Range("A" & CStr(headerRow + 1) & ":M" & CStr(Application.Max(headerRow + 1, targetRow - 1))).Calculate
End Sub
'''

FORM_BRIDGE_CODE = r'''
Private Sub cmdSave_Click()
    Dim patientType As String
    Dim roomValue As String
    Dim patientId As String

    If Not ValidateForm() Then Exit Sub

    patientType = Trim$(cboPatientType.Value)
    If patientType <> "Εξωτερικός" Then
        roomValue = Trim$(cboRoom.Value)
        If Len(roomValue) = 0 Then
            MsgBox "Ο θάλαμος είναι υποχρεωτικός για εσωτερικό ασθενή.", _
                   vbExclamation, "Νέος ασθενής"
            cboRoom.SetFocus
            Exit Sub
        End If
    Else
        roomValue = ""
    End If

    On Error GoTo RegistrationError

    patientId = RegisterPatientInWorkbook( _
        patientType, _
        Trim$(txtHospitalMRN.Text), _
        Trim$(txtDisplayName.Text), _
        roomValue, _
        CBool(chkInfectious.Value), _
        Trim$(cboStatus.Value) _
    )

    MsgBox "Ο ασθενής καταχωρήθηκε επιτυχώς." & vbCrLf & _
           "Patient ID: " & patientId, _
           vbInformation, "Νέος ασθενής"
    Unload Me
    Exit Sub

RegistrationError:
    MsgBox "Η καταχώρηση δεν ολοκληρώθηκε: " & Err.Description, _
           vbCritical, "Νέος ασθενής"
End Sub
'''


class PatientRegistrationBridgeVbaError(RuntimeError):
    """Raised when patient-form bridge wiring cannot be installed safely."""


@dataclass(frozen=True)
class PatientRegistrationBridgePreviewReport:
    source_path: str
    output_path: str
    source_unchanged: bool
    vba_present: bool
    form_present: bool
    bridge_wired: bool


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _workbook_has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def _replace_procedure(code_module, procedure_name: str, replacement: str) -> None:
    try:
        start_line = code_module.ProcStartLine(procedure_name, 0)
        line_count = code_module.ProcCountLines(procedure_name, 0)
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Procedure {procedure_name} was not found in {PATIENT_FORM_NAME}"
        ) from exc
    code_module.DeleteLines(start_line, line_count)
    code_module.InsertLines(start_line, replacement)


def _replace_standard_module(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(BRIDGE_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = BRIDGE_MODULE_NAME
    module.CodeModule.AddFromString(BRIDGE_MODULE_CODE)


def wire_patient_registration_bridge(workbook_path: str | Path) -> None:
    if sys.platform != "win32":
        raise PatientRegistrationBridgeVbaError(
            "Patient registration bridge wiring requires Windows with Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientRegistrationBridgeVbaError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject
        component = vbproject.VBComponents(PATIENT_FORM_NAME)
        code_module = component.CodeModule

        _replace_procedure(code_module, "cmdSave_Click", FORM_BRIDGE_CODE.strip())
        _replace_standard_module(vbproject)

        stored_form = code_module.Lines(1, code_module.CountOfLines)
        stored_module = vbproject.VBComponents(BRIDGE_MODULE_NAME).CodeModule
        stored_module_text = stored_module.Lines(1, stored_module.CountOfLines)
        required_form_markers = (
            "RegisterPatientInWorkbook",
            "Ο ασθενής καταχωρήθηκε επιτυχώς.",
            "Unload Me",
        )
        required_module_markers = (
            "Public Function RegisterPatientInWorkbook",
            "RefreshPlannerRow",
            "RebuildMasterFast",
            "ThisWorkbook.Save",
        )
        if not all(marker in stored_form for marker in required_form_markers):
            raise PatientRegistrationBridgeVbaError(
                "Patient registration fast-path form wiring verification failed before workbook save"
            )
        if not all(marker in stored_module_text for marker in required_module_markers):
            raise PatientRegistrationBridgeVbaError(
                "Patient registration fast-path module verification failed before workbook save"
            )

        workbook.Save()
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Patient registration bridge wiring failed: {exc}"
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def create_patient_registration_bridge_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> PatientRegistrationBridgePreviewReport:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise PatientRegistrationBridgeVbaError(f"Source workbook not found: {source}")
    if source == output:
        raise PatientRegistrationBridgeVbaError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise PatientRegistrationBridgeVbaError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise PatientRegistrationBridgeVbaError(f"Output already exists: {output}")

    before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    try:
        wire_patient_registration_bridge(output)
    except Exception:
        output.unlink(missing_ok=True)
        raise

    after = _sha256(source)
    if before != after:
        output.unlink(missing_ok=True)
        raise PatientRegistrationBridgeVbaError("Source workbook changed during bridge wiring")

    if not _workbook_has_vba(output):
        output.unlink(missing_ok=True)
        raise PatientRegistrationBridgeVbaError("VBA project was not preserved")

    return PatientRegistrationBridgePreviewReport(
        source_path=str(source),
        output_path=str(output),
        source_unchanged=True,
        vba_present=True,
        form_present=True,
        bridge_wired=True,
    )
