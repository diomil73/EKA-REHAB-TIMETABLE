from __future__ import annotations

THERAPIST_FORM_NAME = "frmNewTherapist"

THERAPIST_FORM_CODE = r'''Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Νέος θεραπευτής"
        .Width = 460
        .Height = 260
        .StartUpPosition = 1
        .BackColor = RGB(245, 247, 250)
    End With

    With lblTitle
        .Caption = "Εγγραφή νέου θεραπευτή"
        .Left = 35
        .Top = 20
        .Width = 380
        .Height = 26
        .TextAlign = 2
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Bold = True
        .BackStyle = 0
    End With

    With lblDisplayName
        .Caption = "Ονοματεπώνυμο *"
        .Left = 38
        .Top = 82
        .Width = 135
        .Height = 20
        .Font.Name = "Calibri"
        .Font.Size = 10
        .BackStyle = 0
    End With

    With txtDisplayName
        .Left = 180
        .Top = 78
        .Width = 235
        .Height = 24
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With

    With lblInfo
        .Caption = "Ο θεραπευτής θα προστεθεί με ασφαλές preview στο SETTINGS / THERAPISTS_FTH."
        .Left = 38
        .Top = 120
        .Width = 375
        .Height = 34
        .WordWrap = True
        .Font.Name = "Calibri"
        .Font.Size = 9
        .BackStyle = 0
    End With

    With cmdCancel
        .Caption = "Ακύρωση"
        .Left = 82
        .Top = 175
        .Width = 120
        .Height = 32
        .Cancel = True
    End With

    With cmdSave
        .Caption = "Έλεγχος και preview"
        .Left = 220
        .Top = 175
        .Width = 170
        .Height = 32
        .Default = True
    End With

    txtDisplayName.SetFocus
End Sub

Private Function ValidateForm() As Boolean
    If Len(Trim$(txtDisplayName.Text)) = 0 Then
        MsgBox "Το ονοματεπώνυμο είναι υποχρεωτικό.", vbExclamation, "Νέος θεραπευτής"
        txtDisplayName.SetFocus
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
        MsgBox "Δεν βρέθηκε το registration bridge του συστήματος.", vbCritical, "Νέος θεραπευτής"
        Exit Sub
    End If

    requestPath = Environ$("TEMP") & "\eka_therapist_registration_request_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    responsePath = Environ$("TEMP") & "\eka_therapist_registration_response_" & Format$(Now, "yyyymmdd_hhnnss") & ".json"
    WriteUtf8Text requestPath, BuildTherapistRegistrationJson()

    commandLine = QuoteArg("python") & " " & QuoteArg(bridgeScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo BridgeError
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)

    If Dir$(responsePath) = "" Then
        MsgBox "Το registration backend δεν επέστρεψε αποτέλεσμα.", vbCritical, "Νέος θεραπευτής"
        GoTo CleanUp
    End If

    responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η εγγραφή δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & _
               JsonStringValue(responseText, "error"), vbExclamation, "Νέος θεραπευτής"
        GoTo CleanUp
    End If

    MsgBox "Δημιουργήθηκε ασφαλές preview εγγραφής." & vbCrLf & _
           "Θεραπευτής: " & JsonStringValue(responseText, "display_name") & vbCrLf & _
           "Αρχείο: " & JsonStringValue(responseText, "output_path"), _
           vbInformation, "Νέος θεραπευτής"

    Unload Me

CleanUp:
    On Error Resume Next
    If Len(requestPath) > 0 Then Kill requestPath
    If Len(responsePath) > 0 Then Kill responsePath
    On Error GoTo 0
    Exit Sub

BridgeError:
    MsgBox "Δεν ήταν δυνατή η εκτέλεση του registration backend: " & Err.Description, _
           vbCritical, "Νέος θεραπευτής"
    Resume CleanUp
End Sub

Private Function BuildTherapistRegistrationJson() As String
    Dim q As String
    q = Chr$(34)

    BuildTherapistRegistrationJson = _
        "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_dir" & q & ":" & q & JsonEscape(ThisWorkbook.Path) & q & "," & _
        q & "action" & q & ":" & q & "new_therapist" & q & "," & _
        q & "overwrite" & q & ":true," & _
        q & "values" & q & ":{" & _
        q & "display_name" & q & ":" & q & JsonEscape(Trim$(txtDisplayName.Text)) & q & _
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


def install_therapist_form(vbproject, *, position_control) -> None:
    try:
        existing = vbproject.VBComponents(THERAPIST_FORM_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    form = vbproject.VBComponents.Add(3)
    form.Name = THERAPIST_FORM_NAME
    designer = form.Designer
    designer.Caption = "Νέος θεραπευτής"

    controls = (
        ("Forms.Label.1", "lblTitle", "Εγγραφή νέου θεραπευτή", 20),
        ("Forms.Label.1", "lblDisplayName", "Ονοματεπώνυμο *", 82),
        ("Forms.TextBox.1", "txtDisplayName", None, 78),
        ("Forms.Label.1", "lblInfo", "", 120),
        ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 175),
        ("Forms.CommandButton.1", "cmdSave", "Έλεγχος και preview", 175),
    )

    for prog_id, name, caption, top in controls:
        control = designer.Controls.Add(prog_id, name, True)
        if caption is not None:
            control.Caption = caption
        position_control(control, 24, top)

    form.CodeModule.AddFromString(THERAPIST_FORM_CODE)
