from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

FORM_NAME = "frmOutpatientSchedule"
MODULE_NAME = "modOutpatientSchedule"
MACRO_NAME = "ShowOutpatientScheduleForm"

MODULE_CODE = f'''Option Explicit

Public Sub {MACRO_NAME}()
    {FORM_NAME}.Show
End Sub
'''

FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Πρόγραμμα εξωτερικού ασθενή"
    Me.Width = 560
    Me.Height = 625
    Me.StartUpPosition = 1
    Me.BackColor = RGB(245, 247, 250)

    StyleLabel lblPatient, "Εξωτερικός ασθενής *", 48
    StyleCombo cboPatient, 44
    StyleLabel lblTreatment, "Θεραπεία *", 92
    StyleCombo cboTreatment, 88
    StyleLabel lblTime, "Ώρα *", 136
    StyleCombo cboTime, 132
    StyleLabel lblDays, "Ημέρες *", 180
    StyleCombo cboDays, 176
    StyleLabel lblTherapist, "Θεραπευτής *", 224
    StyleCombo cboTherapist, 220
    StyleLabel lblSuggestMode, "Τι θα θέλατε να σας προτείνω;", 268
    StyleCombo cboSuggestMode, 288

    lblTitle.Caption = "Επαναλαμβανόμενο πρόγραμμα εξωτερικού ασθενή"
    lblTitle.Left = 35
    lblTitle.Top = 15
    lblTitle.Width = 470
    lblTitle.Height = 24
    lblTitle.TextAlign = 2
    lblTitle.Font.Name = "Calibri"
    lblTitle.Font.Size = 14
    lblTitle.Font.Bold = True
    lblTitle.BackStyle = 0

    cmdSuggest.Caption = "Βρες εναλλακτικές"
    cmdSuggest.Left = 190
    cmdSuggest.Top = 323
    cmdSuggest.Width = 180
    cmdSuggest.Height = 30

    lstSuggestions.Left = 38
    lstSuggestions.Top = 365
    lstSuggestions.Width = 475
    lstSuggestions.Height = 95
    lstSuggestions.Font.Name = "Calibri"
    lstSuggestions.Font.Size = 10

    cmdUseSuggestion.Caption = "Χρήση επιλογής"
    cmdUseSuggestion.Left = 190
    cmdUseSuggestion.Top = 470
    cmdUseSuggestion.Width = 180
    cmdUseSuggestion.Height = 28

    lblInfo.Caption = "Οι προτάσεις κρατούν σταθερά όσα δεν επιλέξατε να αλλάξουν και ελέγχονται απέναντι στο υπάρχον πρόγραμμα."
    lblInfo.Left = 38
    lblInfo.Top = 505
    lblInfo.Width = 475
    lblInfo.Height = 32
    lblInfo.WordWrap = True
    lblInfo.BackStyle = 0

    cmdCancel.Caption = "Ακύρωση"
    cmdCancel.Left = 125
    cmdCancel.Top = 550
    cmdCancel.Width = 125
    cmdCancel.Height = 32
    cmdCancel.Cancel = True

    cmdSave.Caption = "Έλεγχος και preview"
    cmdSave.Left = 280
    cmdSave.Top = 550
    cmdSave.Width = 170
    cmdSave.Height = 32
    cmdSave.Default = True

    LoadPatients
    LoadSettings
    LoadSuggestionModes
End Sub

Private Sub StyleLabel(ByVal control As MSForms.Label, ByVal text As String, ByVal topPosition As Single)
    control.Caption = text
    control.Left = 38
    control.Top = topPosition
    control.Width = 145
    If control.Name = "lblSuggestMode" Then control.Width = 230
    control.Height = 20
    control.Font.Name = "Calibri"
    control.Font.Size = 10
    control.BackStyle = 0
End Sub

Private Sub StyleCombo(ByVal control As MSForms.ComboBox, ByVal topPosition As Single)
    control.Left = 190
    control.Top = topPosition
    control.Width = 323
    control.Height = 24
    control.Font.Name = "Calibri"
    control.Font.Size = 10
    control.Style = 2
End Sub

Private Sub LoadSuggestionModes()
    cboSuggestMode.Clear
    cboSuggestMode.AddItem "Αλλαγή ώρας"
    cboSuggestMode.AddItem "Αλλαγή ημερών"
    cboSuggestMode.AddItem "Αλλαγή ώρας και ημερών"
    cboSuggestMode.AddItem "Αλλαγή θεραπευτή"
    cboSuggestMode.AddItem "Αλλαγή θεραπευτή και ώρας"
    cboSuggestMode.AddItem "Αλλαγή θεραπευτή και ημερών"
    cboSuggestMode.AddItem "Αλλαγή θεραπευτή, ώρας και ημερών"
    cboSuggestMode.ListIndex = 0
End Sub

Private Sub LoadPatients()
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim patientType As String
    Dim patientId As String
    Dim patientName As String

    Set ws = ThisWorkbook.Worksheets("PATIENTS")
    cboPatient.Clear
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        patientType = Trim$(CStr(ws.Cells(rowIndex, 6).Value))
        If patientType = "Εξωτερικός" Or LCase$(patientType) = "outpatient" Then
            patientId = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
            patientName = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
            If Len(patientId) > 0 And Len(patientName) > 0 Then
                cboPatient.AddItem patientId & " | " & patientName
            End If
        End If
    Next rowIndex
End Sub

Private Sub LoadSettings()
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim valueText As String

    Set ws = ThisWorkbook.Worksheets("SETTINGS")

    cboTherapist.Clear
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(valueText) > 0 Then cboTherapist.AddItem valueText
    Next r

    cboTime.Clear
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 2).Value) Then
            cboTime.AddItem Format$(ws.Cells(r, 2).Value, "hh:mm")
        Else
            valueText = Trim$(CStr(ws.Cells(r, 2).Text))
            If Len(valueText) > 0 Then cboTime.AddItem valueText
        End If
    Next r

    cboDays.Clear
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    For r = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(r, 4).Value))
        If Len(valueText) > 0 Then cboDays.AddItem valueText
    Next r

    cboTreatment.Clear
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    For r = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(r, 5).Value))
        If Len(valueText) > 0 And valueText <> "Ρομποτικό" Then cboTreatment.AddItem valueText
    Next r
End Sub

Private Function PatientIdFromSelection() As String
    Dim p As Long
    p = InStr(1, cboPatient.Value, " | ")
    If p > 0 Then PatientIdFromSelection = Left$(cboPatient.Value, p - 1)
End Function

Private Function ValidateForm() As Boolean
    If Len(PatientIdFromSelection()) = 0 Then MsgBox "Επιλέξτε εξωτερικό ασθενή.", vbExclamation: Exit Function
    If Len(Trim$(cboTreatment.Value)) = 0 Then MsgBox "Επιλέξτε θεραπεία.", vbExclamation: Exit Function
    If Len(Trim$(cboTime.Value)) = 0 Then MsgBox "Επιλέξτε ώρα.", vbExclamation: Exit Function
    If Len(Trim$(cboDays.Value)) = 0 Then MsgBox "Επιλέξτε ημέρες.", vbExclamation: Exit Function
    If Len(Trim$(cboTherapist.Value)) = 0 Then MsgBox "Επιλέξτε θεραπευτή.", vbExclamation: Exit Function
    ValidateForm = True
End Function

Private Function ValidateSuggestionRequest() As Boolean
    If Len(Trim$(cboTherapist.Value)) = 0 Then MsgBox "Επιλέξτε πρώτα θεραπευτή προτίμησης.", vbExclamation: Exit Function
    If Len(Trim$(cboTime.Value)) = 0 Then MsgBox "Επιλέξτε πρώτα ώρα προτίμησης.", vbExclamation: Exit Function
    If Len(Trim$(cboDays.Value)) = 0 Then MsgBox "Επιλέξτε πρώτα ημέρες προτίμησης.", vbExclamation: Exit Function
    If cboSuggestMode.ListIndex < 0 Then MsgBox "Επιλέξτε τι θέλετε να αλλάξει.", vbExclamation: Exit Function
    ValidateSuggestionRequest = True
End Function

Private Sub cmdSuggest_Click()
    Dim requestPath As String
    Dim responsePath As String
    Dim bridgeScript As String
    Dim commandLine As String
    Dim responseText As String
    Dim linesText As String
    Dim rows() As String
    Dim row As Variant
    Dim exitCode As Long

    If Not ValidateSuggestionRequest() Then Exit Sub

    bridgeScript = ResolveSuggestionBridgeScript()
    If Len(bridgeScript) = 0 Then
        MsgBox "Δεν βρέθηκε το suggestion bridge.", vbCritical
        Exit Sub
    End If

    requestPath = Environ$("TEMP") & "\eka_outpatient_suggestion_request_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    responsePath = Environ$("TEMP") & "\eka_outpatient_suggestion_response_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    WriteUtf8Text requestPath, BuildSuggestionRequestJson()

    commandLine = QuoteArg("python") & " " & QuoteArg(bridgeScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo SuggestionError
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) = "" Then
        MsgBox "Το suggestion backend δεν επέστρεψε αποτέλεσμα.", vbCritical
        GoTo SuggestionCleanUp
    End If

    responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Δεν ήταν δυνατή η εύρεση εναλλακτικών:" & vbCrLf & vbCrLf & JsonStringValue(responseText, "error"), vbExclamation
        GoTo SuggestionCleanUp
    End If

    lstSuggestions.Clear
    linesText = JsonStringValue(responseText, "suggestion_lines")
    If Len(linesText) = 0 Then
        MsgBox "Δεν βρέθηκαν ασφαλείς εναλλακτικές με αυτά τα κριτήρια.", vbInformation
        GoTo SuggestionCleanUp
    End If

    rows = Split(linesText, vbLf)
    For Each row In rows
        If Len(Trim$(CStr(row))) > 0 Then lstSuggestions.AddItem CStr(row)
    Next row

SuggestionCleanUp:
    On Error Resume Next
    Kill requestPath
    Kill responsePath
    On Error GoTo 0
    Exit Sub

SuggestionError:
    MsgBox "Δεν ήταν δυνατή η εκτέλεση του suggestion backend: " & Err.Description, vbCritical
    Resume SuggestionCleanUp
End Sub

Private Sub cmdUseSuggestion_Click()
    ApplySelectedSuggestion
End Sub

Private Sub lstSuggestions_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ApplySelectedSuggestion
End Sub

Private Sub ApplySelectedSuggestion()
    Dim parts() As String
    If lstSuggestions.ListIndex < 0 Then
        MsgBox "Επιλέξτε πρώτα μία πρόταση.", vbExclamation
        Exit Sub
    End If
    parts = Split(CStr(lstSuggestions.Value), " | ")
    If UBound(parts) <> 2 Then
        MsgBox "Η πρόταση δεν έχει αναμενόμενη μορφή.", vbCritical
        Exit Sub
    End If
    cboTherapist.Value = Trim$(parts(0))
    cboTime.Value = Trim$(parts(1))
    cboDays.Value = Trim$(parts(2))
End Sub

Private Sub cmdSave_Click()
    Dim requestPath As String
    Dim responsePath As String
    Dim bridgeScript As String
    Dim commandLine As String
    Dim responseText As String
    Dim exitCode As Long

    If Not ValidateForm() Then Exit Sub

    bridgeScript = ResolveBridgeScript()
    If Len(bridgeScript) = 0 Then
        MsgBox "Δεν βρέθηκε το outpatient schedule bridge.", vbCritical
        Exit Sub
    End If

    requestPath = Environ$("TEMP") & "\eka_outpatient_schedule_request_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    responsePath = Environ$("TEMP") & "\eka_outpatient_schedule_response_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    WriteUtf8Text requestPath, BuildRequestJson()

    commandLine = QuoteArg("python") & " " & QuoteArg(bridgeScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo BridgeError
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    If Dir$(responsePath) = "" Then
        MsgBox "Το backend δεν επέστρεψε αποτέλεσμα.", vbCritical
        GoTo CleanUp
    End If

    responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η καταχώρηση δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & JsonStringValue(responseText, "error"), vbExclamation
        GoTo CleanUp
    End If

    MsgBox "Δημιουργήθηκε ασφαλές preview προγράμματος." & vbCrLf & _
           "Base entry: " & JsonStringValue(responseText, "base_entry_id") & vbCrLf & _
           "Αρχείο: " & JsonStringValue(responseText, "output_path"), vbInformation

CleanUp:
    On Error Resume Next
    Kill requestPath
    Kill responsePath
    On Error GoTo 0
    Exit Sub

BridgeError:
    MsgBox "Δεν ήταν δυνατή η εκτέλεση του backend: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Private Function BuildRequestJson() As String
    Dim q As String
    q = Chr$(34)
    BuildRequestJson = "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_dir" & q & ":" & q & JsonEscape(ThisWorkbook.Path) & q & "," & _
        q & "overwrite" & q & ":true," & _
        q & "values" & q & ":{" & _
        q & "patient_id" & q & ":" & q & JsonEscape(PatientIdFromSelection()) & q & "," & _
        q & "treatment" & q & ":" & q & JsonEscape(cboTreatment.Value) & q & "," & _
        q & "time" & q & ":" & q & JsonEscape(cboTime.Value) & q & "," & _
        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _
        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & _
        "}}"
End Function

Private Function BuildSuggestionRequestJson() As String
    Dim q As String
    Dim modeNumber As Long
    q = Chr$(34)
    modeNumber = cboSuggestMode.ListIndex + 1
    BuildSuggestionRequestJson = "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "values" & q & ":{" & _
        q & "mode" & q & ":" & CStr(modeNumber) & "," & _
        q & "therapist" & q & ":" & q & JsonEscape(cboTherapist.Value) & q & "," & _
        q & "time" & q & ":" & q & JsonEscape(cboTime.Value) & q & "," & _
        q & "days" & q & ":" & q & JsonEscape(cboDays.Value) & q & "," & _
        q & "limit" & q & ":5" & _
        "}}"
End Function

Private Function ResolveBridgeScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\outpatient_schedule_bridge_cli.py"
    If Dir$(candidate) <> "" Then ResolveBridgeScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\outpatient_schedule_bridge_cli.py"
    If Dir$(candidate) <> "" Then ResolveBridgeScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
End Function

Private Function ResolveSuggestionBridgeScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\outpatient_schedule_suggestion_bridge_cli.py"
    If Dir$(candidate) <> "" Then ResolveSuggestionBridgeScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\outpatient_schedule_suggestion_bridge_cli.py"
    If Dir$(candidate) <> "" Then ResolveSuggestionBridgeScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
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
    Dim marker As String, startPos As Long, i As Long, ch As String, escaped As Boolean, value As String
    marker = Chr$(34) & key & Chr$(34) & ":"
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos + Len(marker), jsonText, Chr$(34))
    If startPos = 0 Then Exit Function
    For i = startPos + 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)
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
    Next i
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'''


class OutpatientScheduleFormVbaError(RuntimeError):
    pass


@dataclass(frozen=True)
class OutpatientScheduleFormPreviewReport:
    source_path: str
    output_path: str
    source_unchanged: bool
    vba_present: bool
    form_present: bool
    macro_name: str = MACRO_NAME


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _has_vba(path: Path) -> bool:
    with ZipFile(path) as archive:
        return "xl/vbaProject.bin" in archive.namelist()


def create_outpatient_schedule_form_preview(source_path: str | Path, output_path: str | Path, *, overwrite: bool = False) -> OutpatientScheduleFormPreviewReport:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise OutpatientScheduleFormVbaError(f"Source workbook not found: {source}")
    if source == output:
        raise OutpatientScheduleFormVbaError("Output must differ from source")
    if output.exists() and not overwrite:
        raise OutpatientScheduleFormVbaError(f"Output already exists: {output}")

    before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    if sys.platform != "win32":
        raise OutpatientScheduleFormVbaError("Windows Excel is required")
    import win32com.client  # type: ignore[import-not-found]

    excel = None
    workbook = None
    failure: Exception | None = None
    form_present = False
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject

        for name in (FORM_NAME, MODULE_NAME):
            try:
                vbproject.VBComponents.Remove(vbproject.VBComponents(name))
            except Exception:
                pass

        module = vbproject.VBComponents.Add(1)
        module.Name = MODULE_NAME
        module.CodeModule.AddFromString(MODULE_CODE)

        form = vbproject.VBComponents.Add(3)
        form.Name = FORM_NAME
        form.Designer.Caption = "Πρόγραμμα εξωτερικού ασθενή"
        for prog_id, name, caption, top in (
            ("Forms.Label.1", "lblTitle", "", 15),
            ("Forms.Label.1", "lblPatient", "", 48),
            ("Forms.ComboBox.1", "cboPatient", "", 44),
            ("Forms.Label.1", "lblTreatment", "", 92),
            ("Forms.ComboBox.1", "cboTreatment", "", 88),
            ("Forms.Label.1", "lblTime", "", 136),
            ("Forms.ComboBox.1", "cboTime", "", 132),
            ("Forms.Label.1", "lblDays", "", 180),
            ("Forms.ComboBox.1", "cboDays", "", 176),
            ("Forms.Label.1", "lblTherapist", "", 224),
            ("Forms.ComboBox.1", "cboTherapist", "", 220),
            ("Forms.Label.1", "lblSuggestMode", "", 268),
            ("Forms.ComboBox.1", "cboSuggestMode", "", 288),
            ("Forms.CommandButton.1", "cmdSuggest", "Βρες εναλλακτικές", 323),
            ("Forms.ListBox.1", "lstSuggestions", "", 365),
            ("Forms.CommandButton.1", "cmdUseSuggestion", "Χρήση επιλογής", 470),
            ("Forms.Label.1", "lblInfo", "", 505),
            ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 550),
            ("Forms.CommandButton.1", "cmdSave", "Έλεγχος και preview", 550),
        ):
            control = form.Designer.Controls.Add(prog_id, name, True)
            if caption:
                control.Caption = caption
            control.Left = 24
            control.Top = top
        form.CodeModule.AddFromString(FORM_CODE)
        workbook.Save()

        form_present = True
        try:
            _ = vbproject.VBComponents(FORM_NAME)
        except Exception:
            form_present = False
    except Exception as exc:
        failure = exc
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

    if failure is not None:
        output.unlink(missing_ok=True)
        raise OutpatientScheduleFormVbaError(str(failure)) from failure

    if _sha256(source) != before:
        output.unlink(missing_ok=True)
        raise OutpatientScheduleFormVbaError("Source workbook changed")
    if not _has_vba(output) or not form_present:
        output.unlink(missing_ok=True)
        raise OutpatientScheduleFormVbaError("Form/VBA verification failed")

    return OutpatientScheduleFormPreviewReport(str(source), str(output), True, True, True)
