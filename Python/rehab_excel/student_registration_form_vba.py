from __future__ import annotations

STUDENT_FORM_NAME = "frmNewStudent"

STUDENT_FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Νέος φοιτητής"
        .Width = 520
        .Height = 500
        .StartUpPosition = 1
        .BackColor = RGB(245, 247, 250)
    End With

    StyleTitle lblTitle, "Εγγραφή νέου φοιτητή", 18
    StyleLabel lblStudentID, "Student ID *", 62
    StyleTextBox txtStudentID, 58
    StyleLabel lblDisplayName, "Ονοματεπώνυμο *", 106
    StyleTextBox txtDisplayName, 102
    StyleLabel lblPlacementStart, "Έναρξη πρακτικής *", 150
    StyleTextBox txtPlacementStart, 146
    StyleLabel lblPlacementEnd, "Λήξη πρακτικής *", 194
    StyleTextBox txtPlacementEnd, 190
    StyleLabel lblSupervisor, "Επόπτης θεραπευτής", 238
    StyleComboBox cboSupervisor, 234
    StyleLabel lblReplacement, "Αναπληρώσεις", 282
    StyleCheckBox chkReplacement, "Ναι", 278, True
    StyleLabel lblRobotic, "Ρομποτική αποκατάσταση", 322
    StyleCheckBox chkRobotic, "Ναι", 318, False

    With lblInfo
        .Caption = "Ημερομηνίες: ΗΗ/ΜΜ/ΕΕΕΕ. Ο επόπτης, αν δηλωθεί, πρέπει να υπάρχει στους θεραπευτές."
        .Left = 38
        .Top = 356
        .Width = 440
        .Height = 32
        .WordWrap = True
        .Font.Name = "Calibri"
        .Font.Size = 9
        .ForeColor = RGB(71, 85, 105)
        .BackStyle = 0
    End With

    With cmdCancel
        .Caption = "Ακύρωση"
        .Left = 105
        .Top = 408
        .Width = 125
        .Height = 34
        .Cancel = True
    End With

    With cmdSave
        .Caption = "Έλεγχος και preview"
        .Left = 250
        .Top = 408
        .Width = 175
        .Height = 34
        .Default = True
    End With

    LoadTherapists
    txtStudentID.SetFocus
End Sub

Private Sub StyleTitle(ByVal control As MSForms.Label, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 38
        .Top = topPosition
        .Width = 440
        .Height = 26
        .TextAlign = 2
        .Font.Name = "Calibri"
        .Font.Size = 15
        .Font.Bold = True
        .ForeColor = RGB(45, 55, 72)
        .BackStyle = 0
    End With
End Sub

Private Sub StyleLabel(ByVal control As MSForms.Label, ByVal text As String, ByVal topPosition As Single)
    With control
        .Caption = text
        .Left = 38
        .Top = topPosition
        .Width = 175
        .Height = 20
        .Font.Name = "Calibri"
        .Font.Size = 10
        .ForeColor = RGB(51, 65, 85)
        .BackStyle = 0
    End With
End Sub

Private Sub StyleTextBox(ByVal control As MSForms.TextBox, ByVal topPosition As Single)
    With control
        .Left = 220
        .Top = topPosition
        .Width = 255
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub

Private Sub StyleComboBox(ByVal control As MSForms.ComboBox, ByVal topPosition As Single)
    With control
        .Left = 220
        .Top = topPosition
        .Width = 255
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Style = 2
    End With
End Sub

Private Sub StyleCheckBox(ByVal control As MSForms.CheckBox, ByVal text As String, ByVal topPosition As Single, ByVal defaultValue As Boolean)
    With control
        .Caption = text
        .Left = 220
        .Top = topPosition
        .Width = 90
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Value = defaultValue
    End With
End Sub

Private Sub LoadTherapists()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim valueText As String

    On Error GoTo SettingsError
    Set ws = ThisWorkbook.Worksheets("SETTINGS")
    cboSupervisor.Clear
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        valueText = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
        If Len(valueText) > 0 Then cboSupervisor.AddItem valueText
    Next rowIndex
    Exit Sub

SettingsError:
    MsgBox "Δεν ήταν δυνατή η φόρτωση των θεραπευτών από το SETTINGS.", vbExclamation, "Νέος φοιτητής"
End Sub

Private Function ValidateForm() As Boolean
    If Len(Trim$(txtStudentID.Text)) = 0 Then
        MsgBox "Το Student ID είναι υποχρεωτικό.", vbExclamation, "Νέος φοιτητής"
        txtStudentID.SetFocus
        Exit Function
    End If
    If Len(Trim$(txtDisplayName.Text)) = 0 Then
        MsgBox "Το ονοματεπώνυμο είναι υποχρεωτικό.", vbExclamation, "Νέος φοιτητής"
        txtDisplayName.SetFocus
        Exit Function
    End If
    If Not IsDate(txtPlacementStart.Text) Then
        MsgBox "Η ημερομηνία έναρξης δεν είναι έγκυρη.", vbExclamation, "Νέος φοιτητής"
        txtPlacementStart.SetFocus
        Exit Function
    End If
    If Not IsDate(txtPlacementEnd.Text) Then
        MsgBox "Η ημερομηνία λήξης δεν είναι έγκυρη.", vbExclamation, "Νέος φοιτητής"
        txtPlacementEnd.SetFocus
        Exit Function
    End If
    If CDate(txtPlacementEnd.Text) < CDate(txtPlacementStart.Text) Then
        MsgBox "Η λήξη πρακτικής δεν μπορεί να είναι πριν από την έναρξη.", vbExclamation, "Νέος φοιτητής"
        txtPlacementEnd.SetFocus
        Exit Function
    End If
    ValidateForm = True
End Function

Private Sub cmdSave_Click()
    Dim requestPath As String
    Dim responsePath As String
    Dim bridgeScript As String
    Dim commandLine As String
    Dim responseText As String
    Dim exitCode As Long

    If Not ValidateForm() Then Exit Sub

    bridgeScript = ResolveRegistrationBridgeScript()
    If Len(bridgeScript) = 0 Then
        MsgBox "Δεν βρέθηκε το registration bridge του συστήματος.", vbCritical, "Νέος φοιτητής"
        Exit Sub
    End If

    requestPath = Environ$("TEMP") & "\eka_student_registration_request_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    responsePath = Environ$("TEMP") & "\eka_student_registration_response_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    WriteUtf8Text requestPath, BuildStudentRegistrationJson()

    commandLine = QuoteArg("python") & " " & QuoteArg(bridgeScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo BridgeError
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)

    If Dir$(responsePath) = "" Then
        MsgBox "Το registration backend δεν επέστρεψε αποτέλεσμα.", vbCritical, "Νέος φοιτητής"
        GoTo CleanUp
    End If

    responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η εγγραφή δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & _
               JsonStringValue(responseText, "error"), vbExclamation, "Νέος φοιτητής"
        GoTo CleanUp
    End If

    MsgBox "Δημιουργήθηκε ασφαλές preview εγγραφής." & vbCrLf & _
           "Student ID: " & JsonStringValue(responseText, "subject_key") & vbCrLf & _
           "Αρχείο: " & JsonStringValue(responseText, "output_path"), _
           vbInformation, "Νέος φοιτητής"
    Unload Me

CleanUp:
    On Error Resume Next
    If Len(requestPath) > 0 Then Kill requestPath
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0
    Exit Sub

BridgeError:
    MsgBox "Δεν ήταν δυνατή η εκτέλεση του registration backend: " & Err.Description, _
           vbCritical, "Νέος φοιτητής"
    Resume CleanUp
End Sub

Private Function BuildStudentRegistrationJson() As String
    Dim q As String
    Dim supervisorValue As String
    q = Chr$(34)
    supervisorValue = Trim$(cboSupervisor.Value)

    BuildStudentRegistrationJson = _
        "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_dir" & q & ":" & q & JsonEscape(ThisWorkbook.Path) & q & "," & _
        q & "action" & q & ":" & q & "new_student" & q & "," & _
        q & "overwrite" & q & ":true," & _
        q & "values" & q & ":{" & _
        q & "student_id" & q & ":" & q & JsonEscape(Trim$(txtStudentID.Text)) & q & "," & _
        q & "display_name" & q & ":" & q & JsonEscape(Trim$(txtDisplayName.Text)) & q & "," & _
        q & "placement_start" & q & ":" & q & Format$(CDate(txtPlacementStart.Text), "dd/mm/yyyy") & q & "," & _
        q & "placement_end" & q & ":" & q & Format$(CDate(txtPlacementEnd.Text), "dd/mm/yyyy") & q & "," & _
        q & "supervisor_therapist_id" & q & ":" & q & JsonEscape(supervisorValue) & q & "," & _
        q & "replacement_capable" & q & ":" & LCase$(CStr(chkReplacement.Value)) & "," & _
        q & "robotic_capable" & q & ":" & LCase$(CStr(chkRobotic.Value)) & _
        "}}"
End Function

Private Function ResolveRegistrationBridgeScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\Python\tools\registration_bridge_cli.py"
    If Dir$(candidate) <> "" Then ResolveRegistrationBridgeScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\..\..\Python\tools\registration_bridge_cli.py"
    If Dir$(candidate) <> "" Then ResolveRegistrationBridgeScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
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
    Dim marker As String, startPos As Long, index As Long, ch As String, escaped As Boolean, value As String
    marker = Chr$(34) & key & Chr$(34) & ":"
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos + Len(marker), jsonText, Chr$(34))
    If startPos = 0 Then Exit Function
    For index = startPos + 1 To Len(jsonText)
        ch = Mid$(jsonText, index, 1)
        If escaped Then
            If ch = "n" Then value = value & vbLf Else value = value & ch
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'''


def install_student_form(vbproject, *, position_control) -> None:
    try:
        existing = vbproject.VBComponents(STUDENT_FORM_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    form = vbproject.VBComponents.Add(3)
    form.Name = STUDENT_FORM_NAME
    designer = form.Designer
    designer.Caption = "Νέος φοιτητής"

    controls = (
        ("Forms.Label.1", "lblTitle", "Εγγραφή νέου φοιτητή", 18),
        ("Forms.Label.1", "lblStudentID", "Student ID *", 62),
        ("Forms.TextBox.1", "txtStudentID", None, 58),
        ("Forms.Label.1", "lblDisplayName", "Ονοματεπώνυμο *", 106),
        ("Forms.TextBox.1", "txtDisplayName", None, 102),
        ("Forms.Label.1", "lblPlacementStart", "Έναρξη πρακτικής *", 150),
        ("Forms.TextBox.1", "txtPlacementStart", None, 146),
        ("Forms.Label.1", "lblPlacementEnd", "Λήξη πρακτικής *", 194),
        ("Forms.TextBox.1", "txtPlacementEnd", None, 190),
        ("Forms.Label.1", "lblSupervisor", "Επόπτης θεραπευτής", 238),
        ("Forms.ComboBox.1", "cboSupervisor", None, 234),
        ("Forms.Label.1", "lblReplacement", "Αναπληρώσεις", 282),
        ("Forms.CheckBox.1", "chkReplacement", "Ναι", 278),
        ("Forms.Label.1", "lblRobotic", "Ρομποτική αποκατάσταση", 322),
        ("Forms.CheckBox.1", "chkRobotic", "Ναι", 318),
        ("Forms.Label.1", "lblInfo", "", 356),
        ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 408),
        ("Forms.CommandButton.1", "cmdSave", "Έλεγχος και preview", 408),
    )

    for prog_id, name, caption, top in controls:
        control = designer.Controls.Add(prog_id, name, True)
        if caption is not None:
            control.Caption = caption
        position_control(control, 24, top)

    form.CodeModule.AddFromString(STUDENT_FORM_CODE)
