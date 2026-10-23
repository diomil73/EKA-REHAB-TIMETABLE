from __future__ import annotations

import sys
from pathlib import Path


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
End Sub

Public Function ApplySelectedDoubleBooking(ByVal targetSessionId As String, ByVal therapistId As String, ByVal replacementTime As String) As Boolean
    Dim applyScript As String, outputPath As String, responsePath As String
    Dim commandLine As String, exitCode As Long, responseText As String
    Dim sourceHash As String, previewPath As String

    applyScript = ResolveProjectTool("apply_confirmed_double_booking.py")
    If Len(applyScript) = 0 Then
        MsgBox "Δεν βρέθηκε το εργαλείο εφαρμογής διπλοκράτησης.", vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    outputPath = ThisWorkbook.Path & "\\DOUBLE_BOOKING_CONFIRMED_UI_PREVIEW.xlsm"
    responsePath = Environ$("TEMP") & "\\eka_double_booking_apply_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(applyScript) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --review " & QuoteArg(mReviewPath) & _
                  " --target-session-id " & QuoteArg(targetSessionId) & _
                  " --therapist-id " & QuoteArg(therapistId) & _
                  " --time " & QuoteArg(replacementTime) & _
                  " --output " & QuoteArg(outputPath) & _
                  " --response " & QuoteArg(responsePath) & _
                  " --confirm --overwrite"

    Application.StatusBar = "Εφαρμογή και επαλήθευση διπλοκράτησης..."
    exitCode = RunHidden(commandLine)
    Application.StatusBar = False

    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Then
        MsgBox "Η διπλοκράτηση δεν εφαρμόστηκε." & vbCrLf & vbCrLf & _
               JsonStringValue(responseText, "error"), vbExclamation, "Διπλοκράτηση"
        GoTo CleanUp
    End If

    previewPath = JsonStringValue(responseText, "output_path")
    sourceHash = JsonStringValue(responseText, "source_sha256_before")
    If Len(previewPath) = 0 Or Len(sourceHash) = 0 Then
        MsgBox "Το backend δεν επέστρεψε τα στοιχεία ασφαλούς αποθήκευσης.", vbCritical, "Διπλοκράτηση"
        GoTo CleanUp
    End If

    If Not StartAuthoritativeCommit(previewPath, sourceHash) Then GoTo CleanUp

    MsgBox "Η διπλοκράτηση επαληθεύτηκε και θα αποθηκευτεί με ασφάλεια." & vbCrLf & vbCrLf & _
           "Το αρχείο θα κλείσει προσωρινά και θα ανοίξει ξανά αυτόματα.", _
           vbInformation, "Διπλοκράτηση"

    ApplySelectedDoubleBooking = True
    Unload frmDoubleBookingReview
    On Error Resume Next
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
    ThisWorkbook.Close SaveChanges:=False
    Exit Function

CleanUp:
    On Error Resume Next
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
End Function

Private Function StartAuthoritativeCommit(ByVal previewPath As String, ByVal sourceSha256 As String) As Boolean
    Dim workerScript As String
    Dim requestPath As String, responsePath As String
    Dim commandLine As String

    workerScript = ResolveProjectTool("authoritative_commit_worker_cli.py")
    If Len(workerScript) = 0 Then
        MsgBox "Δεν βρέθηκε το authoritative save worker του συστήματος.", vbCritical, "Διπλοκράτηση"
        Exit Function
    End If

    requestPath = Environ$("TEMP") & "\\eka_double_booking_commit_request_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"
    responsePath = Environ$("TEMP") & "\\eka_double_booking_commit_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    WriteUtf8Text requestPath, BuildAuthoritativeCommitJson(previewPath, sourceSha256)
    commandLine = QuoteArg("python") & " " & QuoteArg(workerScript) & _
                  " --request " & QuoteArg(requestPath) & _
                  " --response " & QuoteArg(responsePath)

    On Error GoTo WorkerError
    CreateObject("WScript.Shell").Run commandLine, 0, False
    StartAuthoritativeCommit = True
    Exit Function

WorkerError:
    MsgBox "Δεν ήταν δυνατή η εκκίνηση της ασφαλούς αποθήκευσης: " & Err.Description, _
           vbCritical, "Διπλοκράτηση"
End Function

Private Function BuildAuthoritativeCommitJson(ByVal previewPath As String, ByVal sourceSha256 As String) As String
    Dim q As String
    q = Chr$(34)
    BuildAuthoritativeCommitJson = _
        "{" & _
        q & "source_path" & q & ":" & q & JsonEscape(ThisWorkbook.FullName) & q & "," & _
        q & "preview_path" & q & ":" & q & JsonEscape(previewPath) & q & "," & _
        q & "expected_source_sha256" & q & ":" & q & JsonEscape(sourceSha256) & q & "," & _
        q & "remove_preview_after_success" & q & ":true," & _
        q & "reopen" & q & ":true," & _
        q & "timeout_seconds" & q & ":30," & _
        q & "poll_seconds" & q & ":0.5," & _
        q & "post_reopen_target" & q & ":" & q & "REPLACEMENTS" & q & _
        "}"
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

    Call modDoubleBookingReview.ApplySelectedDoubleBooking( _
        lstOptions.List(idx, 0), _
        lstOptions.List(idx, 3), _
        lstOptions.List(idx, 5))
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
