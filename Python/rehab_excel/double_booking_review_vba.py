from __future__ import annotations


DOUBLE_BOOKING_FORM_NAME = "frmDoubleBookingReview"
DOUBLE_BOOKING_MODULE_NAME = "modDoubleBookingReview"
DOUBLE_BOOKING_MACRO_NAME = "ShowDoubleBookingReview"


MODULE_CODE = r'''Option Explicit

Private mReviewPath As String
Private mRowsPath As String

Public Sub ShowDoubleBookingReview()
    Dim prepareScript As String
    Dim commandLine As String, exitCode As Long

    prepareScript = ResolveProjectTool("prepare_double_booking_review_ui.py")
    If Len(prepareScript) = 0 Then
        MsgBox "Δεν βρέθηκε το εργαλείο ελέγχου διπλοκράτησης.", vbCritical, "Διπλοκράτηση"
        Exit Sub
    End If

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση.", vbExclamation, "Διπλοκράτηση"
        Exit Sub
    End If

    On Error GoTo SaveError
    If Not ThisWorkbook.Saved Then ThisWorkbook.Save
    On Error GoTo 0

    mReviewPath = Environ$("TEMP") & "\\eka_double_booking_review.json"
    mRowsPath = Environ$("TEMP") & "\\eka_double_booking_rows.tsv"

    On Error Resume Next
    If Dir$(mReviewPath) <> "" Then Kill mReviewPath
    If Dir$(mRowsPath) <> "" Then Kill mRowsPath
    On Error GoTo 0

    Application.StatusBar = "Έλεγχος εκκρεμών συνεδριών και προετοιμασία επιλογών..."
    commandLine = QuoteArg("python") & " " & QuoteArg(prepareScript) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --review " & QuoteArg(mReviewPath) & _
                  " --rows " & QuoteArg(mRowsPath) & " --overwrite"
    exitCode = RunHidden(commandLine)
    Application.StatusBar = False
    If exitCode <> 0 Or Dir$(mReviewPath) = "" Or Dir$(mRowsPath) = "" Then
        MsgBox "Δεν υπάρχουν εκκρεμείς συνεδρίες για ελεγχόμενη διπλοκράτηση ή ο έλεγχος απέτυχε.", _
               vbInformation, "Διπλοκράτηση"
        Exit Sub
    End If

    frmDoubleBookingReview.LoadRows mRowsPath
    frmDoubleBookingReview.Show vbModal
    Exit Sub

SaveError:
    Application.StatusBar = False
    MsgBox "Δεν ήταν δυνατή η ασφαλής αποθήκευση της τρέχουσας κατάστασης: " & Err.Description, _
           vbCritical, "Διπλοκράτηση"
End Sub

Public Function ApplySelectedDoubleBooking(ByVal targetSessionId As String, ByVal therapistId As String, ByVal replacementTime As String) As Boolean
    Dim applyScript As String, responsePath As String
    Dim commandLine As String, exitCode As Long, responseText As String

    applyScript = ResolveProjectTool("apply_confirmed_double_booking.py")
    If Len(applyScript) = 0 Then
        MsgBox "Δεν βρέθηκε το εργαλείο εφαρμογής διπλοκράτησης.", vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    responsePath = Environ$("TEMP") & "\\eka_double_booking_apply_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(applyScript) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --review " & QuoteArg(mReviewPath) & _
                  " --target-session-id " & QuoteArg(targetSessionId) & _
                  " --therapist-id " & QuoteArg(therapistId) & _
                  " --time " & QuoteArg(replacementTime) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --validate-only --confirm"

    Application.StatusBar = "Επαλήθευση επιβεβαιωμένης διπλοκράτησης..."
    exitCode = RunHidden(commandLine)
    Application.StatusBar = False

    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η διπλοκράτηση δεν εφαρμόστηκε." & vbCrLf & vbCrLf & _
               JsonStringValue(responseText, "error"), vbExclamation, "Διπλοκράτηση"
        GoTo CleanUp
    End If

    If Not PersistValidatedAssignment(responseText) Then GoTo CleanUp

    MsgBox "Η διπλοκράτηση επιβεβαιώθηκε και αποθηκεύτηκε επιτυχώς.", _
           vbInformation, "Διπλοκράτηση"
    ApplySelectedDoubleBooking = True

CleanUp:
    On Error Resume Next
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
End Function

Private Function PersistValidatedAssignment(ByVal responseText As String) As Boolean
    Const FACT_SHEET As String = "DAILY_REPLACEMENT_FACTS"
    Dim ws As Worksheet
    Dim replacementId As String, targetId As String, patientId As String
    Dim originalTherapist As String, replacementTherapist As String
    Dim replacementDate As String, replacementTime As String
    Dim reason As String, providerKind As String, statusText As String
    Dim lastRow As Long, row As Long, writeRow As Long

    replacementId = JsonStringValue(responseText, "replacement_id")
    targetId = JsonStringValue(responseText, "target_session_id")
    patientId = JsonStringValue(responseText, "patient_id")
    originalTherapist = JsonStringValue(responseText, "original_therapist_id")
    replacementTherapist = JsonStringValue(responseText, "replacement_therapist_id")
    replacementDate = JsonStringValue(responseText, "replacement_date")
    replacementTime = JsonStringValue(responseText, "replacement_time")
    reason = JsonStringValue(responseText, "reason")
    providerKind = JsonStringValue(responseText, "provider_kind")
    statusText = JsonStringValue(responseText, "status")

    If Len(replacementId) = 0 Or Len(targetId) = 0 Or Len(patientId) = 0 Or _
       Len(originalTherapist) = 0 Or Len(replacementTherapist) = 0 Or _
       Len(replacementDate) = 0 Or Len(replacementTime) = 0 Or _
       Len(providerKind) = 0 Or UCase$(statusText) <> "ACTIVE" Then
        MsgBox "Το backend δεν επέστρεψε πλήρη επαληθευμένα στοιχεία διπλοκράτησης.", _
               vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FACT_SHEET)
    On Error GoTo PersistError

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = FACT_SHEET
        WriteFactHeaders ws
    ElseIf Not FactHeadersAreValid(ws) Then
        MsgBox "Η δομή του ιστορικού αντικαταστάσεων δεν είναι έγκυρη. Δεν έγινε καμία αλλαγή.", _
               vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    ws.Visible = xlSheetVeryHidden
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    For row = 2 To lastRow
        If Trim$(CStr(ws.Cells(row, 1).Value)) = replacementId Then
            MsgBox "Υπάρχει ήδη καταχώριση με το ίδιο ReplacementId. Δεν έγινε καμία αλλαγή.", _
                   vbCritical, "Διπλοκράτηση"
            Exit Function
        End If
        If Trim$(CStr(ws.Cells(row, 2).Value)) = targetId And _
           NormalizedDateText(ws.Cells(row, 6).Value) = replacementDate And _
           UCase$(Trim$(CStr(ws.Cells(row, 10).Value))) = "ACTIVE" Then
            MsgBox "Η συνεδρία έχει ήδη ενεργή αντικατάσταση. Δεν έγινε καμία αλλαγή.", _
                   vbCritical, "Διπλοκράτηση"
            Exit Function
        End If
    Next row

    writeRow = IIf(lastRow < 2, 2, lastRow + 1)
    ws.Cells(writeRow, 1).Value = replacementId
    ws.Cells(writeRow, 2).Value = targetId
    ws.Cells(writeRow, 3).Value = patientId
    ws.Cells(writeRow, 4).Value = originalTherapist
    ws.Cells(writeRow, 5).Value = replacementTherapist
    ws.Cells(writeRow, 6).NumberFormat = "@"
    ws.Cells(writeRow, 6).Value = replacementDate
    ws.Cells(writeRow, 7).NumberFormat = "@"
    ws.Cells(writeRow, 7).Value = replacementTime
    ws.Cells(writeRow, 8).Value = reason
    ws.Cells(writeRow, 9).Value = providerKind
    ws.Cells(writeRow, 10).Value = "ACTIVE"

    If Trim$(CStr(ws.Cells(writeRow, 1).Value)) <> replacementId Or _
       Trim$(CStr(ws.Cells(writeRow, 2).Value)) <> targetId Or _
       Trim$(CStr(ws.Cells(writeRow, 5).Value)) <> replacementTherapist Or _
       Trim$(CStr(ws.Cells(writeRow, 7).Value)) <> replacementTime Or _
       UCase$(Trim$(CStr(ws.Cells(writeRow, 10).Value))) <> "ACTIVE" Then
        ws.Rows(writeRow).ClearContents
        MsgBox "Η εγγραφή της διπλοκράτησης δεν επαληθεύτηκε. Δεν έγινε αποθήκευση.", _
               vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    ThisWorkbook.Save
    PersistValidatedAssignment = True
    Exit Function

PersistError:
    MsgBox "Δεν ήταν δυνατή η αποθήκευση της διπλοκράτησης: " & Err.Description, _
           vbCritical, "Διπλοκράτηση"
End Function

Private Sub WriteFactHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "ReplacementId"
    ws.Cells(1, 2).Value = "TargetSessionId"
    ws.Cells(1, 3).Value = "PatientId"
    ws.Cells(1, 4).Value = "OriginalTherapistId"
    ws.Cells(1, 5).Value = "ReplacementTherapistId"
    ws.Cells(1, 6).Value = "ReplacementDate"
    ws.Cells(1, 7).Value = "ReplacementTime"
    ws.Cells(1, 8).Value = "Reason"
    ws.Cells(1, 9).Value = "ProviderKind"
    ws.Cells(1, 10).Value = "Status"
End Sub

Private Function FactHeadersAreValid(ByVal ws As Worksheet) As Boolean
    Dim headers As Variant, i As Long
    headers = Array("ReplacementId", "TargetSessionId", "PatientId", "OriginalTherapistId", _
                    "ReplacementTherapistId", "ReplacementDate", "ReplacementTime", _
                    "Reason", "ProviderKind", "Status")
    For i = 0 To UBound(headers)
        If Trim$(CStr(ws.Cells(1, i + 1).Value)) <> CStr(headers(i)) Then Exit Function
    Next i
    FactHeadersAreValid = True
End Function

Private Function NormalizedDateText(ByVal value As Variant) As String
    If IsDate(value) Then
        NormalizedDateText = Format$(CDate(value), "yyyy-mm-dd")
    Else
        NormalizedDateText = Trim$(CStr(value))
    End If
End Function

Private Function RunHidden(ByVal commandLine As String) As Long
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    RunHidden = shell.Run(commandLine, 0, True)
End Function

Private Function ResolveProjectTool(ByVal fileName As String) As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then
        ResolveProjectTool = candidate
        Exit Function
    End If
    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\" & fileName
    If Dir$(candidate) <> "" Then
        ResolveProjectTool = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
    End If
End Function

Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String) As String
    Dim marker As String, startPos As Long, index As Long
    Dim ch As String, escaped As Boolean, value As String
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


FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Ελεγχόμενη διπλοκράτηση"
    Me.Width = 620
    Me.Height = 390
    Me.StartUpPosition = 1
    Me.BackColor = RGB(245, 247, 250)

    lblTitle.Caption = "Επιλογή εξαίρεσης για εκκρεμή συνεδρία"
    lblTitle.Font.Name = "Calibri"
    lblTitle.Font.Size = 14
    lblTitle.Font.Bold = True
    lblTitle.ForeColor = RGB(45, 55, 72)

    lstOptions.ColumnCount = 7
    lstOptions.ColumnWidths = "0 pt;150 pt;60 pt;0 pt;150 pt;60 pt;0 pt"
    lstOptions.Font.Name = "Calibri"
    lstOptions.Font.Size = 10

    lblWarning.Caption = ""
    lblWarning.WordWrap = True
    lblWarning.ForeColor = RGB(160, 80, 0)
    lblWarning.Font.Name = "Calibri"
    lblWarning.Font.Size = 10

    cmdConfirm.Caption = "Επιβεβαίωση διπλοκράτησης"
    cmdConfirm.Font.Bold = True
    cmdCancel.Caption = "Άκυρο"
    cmdCancel.Cancel = True
End Sub

Public Sub LoadRows(ByVal rowsPath As String)
    Dim stream As Object, text As String, lines() As String
    Dim fields() As String, i As Long

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile rowsPath
    text = stream.ReadText
    stream.Close

    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    lstOptions.Clear
    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            fields = Split(lines(i), vbTab)
            If UBound(fields) >= 6 Then
                lstOptions.AddItem fields(0)
                lstOptions.List(lstOptions.ListCount - 1, 1) = fields(1)
                lstOptions.List(lstOptions.ListCount - 1, 2) = fields(2)
                lstOptions.List(lstOptions.ListCount - 1, 3) = fields(3)
                lstOptions.List(lstOptions.ListCount - 1, 4) = fields(4)
                lstOptions.List(lstOptions.ListCount - 1, 5) = fields(5)
                lstOptions.List(lstOptions.ListCount - 1, 6) = fields(6)
            End If
        End If
    Next i
    If lstOptions.ListCount > 0 Then lstOptions.ListIndex = 0
End Sub

Private Sub lstOptions_Change()
    If lstOptions.ListIndex >= 0 Then
        lblWarning.Caption = lstOptions.List(lstOptions.ListIndex, 6)
    Else
        lblWarning.Caption = ""
    End If
End Sub

Private Sub cmdConfirm_Click()
    Dim idx As Long
    If lstOptions.ListIndex < 0 Then
        MsgBox "Επιλέξτε θεραπευτή και ώρα.", vbExclamation, "Διπλοκράτηση"
        Exit Sub
    End If
    idx = lstOptions.ListIndex

    If MsgBox(lblWarning.Caption & vbCrLf & vbCrLf & _
              "Ασθενής: " & lstOptions.List(idx, 1) & vbCrLf & _
              "Αρχική ώρα: " & lstOptions.List(idx, 2) & vbCrLf & _
              "Θεραπευτής: " & lstOptions.List(idx, 4) & vbCrLf & _
              "Ώρα εξαίρεσης: " & lstOptions.List(idx, 5) & vbCrLf & vbCrLf & _
              "Θέλετε να συνεχίσετε;", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Επιβεβαίωση διπλοκράτησης") <> vbYes Then
        Exit Sub
    End If

    If modDoubleBookingReview.ApplySelectedDoubleBooking( _
        lstOptions.List(idx, 0), _
        lstOptions.List(idx, 3), _
        lstOptions.List(idx, 5)) Then
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'''


def install_double_booking_review(vbproject, *, position_control=None) -> None:
    for name in (DOUBLE_BOOKING_MODULE_NAME, DOUBLE_BOOKING_FORM_NAME):
        try:
            existing = vbproject.VBComponents(name)
        except Exception:
            existing = None
        if existing is not None:
            vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = DOUBLE_BOOKING_MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)

    form = vbproject.VBComponents.Add(3)
    form.Name = DOUBLE_BOOKING_FORM_NAME
    designer = form.Designer
    designer.Caption = "Ελεγχόμενη διπλοκράτηση"

    controls = [
        ("Forms.Label.1", "lblTitle", 20, 18, 560, 24),
        ("Forms.ListBox.1", "lstOptions", 20, 55, 560, 150),
        ("Forms.Label.1", "lblWarning", 20, 215, 560, 70),
        ("Forms.CommandButton.1", "cmdConfirm", 270, 305, 190, 32),
        ("Forms.CommandButton.1", "cmdCancel", 470, 305, 110, 32),
    ]
    for prog_id, name, left, top, width, height in controls:
        control = designer.Controls.Add(prog_id, name, True)
        control.Left = left
        control.Top = top
        control.Width = width
        control.Height = height

    form.CodeModule.AddFromString(FORM_CODE)
