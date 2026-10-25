from __future__ import annotations

from .double_booking_review_vba import install_double_booking_review
from .global_replacement_action_vba import install_global_replacement_action
from .therapist_workload_strip_vba import (
    MODULE_CODE as WORKLOAD_MODULE_CODE,
    MODULE_NAME as WORKLOAD_MODULE_NAME,
)


DAILY_INPUT_MODULE_NAME = "modDailyInputAction"
DAILY_INPUT_MACRO_NAME = "ApplyDailyInputPreview"


DAILY_INPUT_MODULE_CODE = r'''Option Explicit

Public Sub ApplyDailyInputPreview()
    Dim scriptPath As String, payloadPath As String, responsePath As String
    Dim commandLine As String, responseText As String
    Dim exitCode As Long, replacementsNeeded As Long

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Ανοίξτε το κανονικά και δοκιμάστε ξανά.", vbExclamation, "Ημερήσια κατάσταση"
        Exit Sub
    End If

    scriptPath = ResolveDailyInputPreparationScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Η λειτουργία εφαρμογής της ημερήσιας κατάστασης δεν είναι διαθέσιμη. Επικοινωνήστε με τον διαχειριστή.", vbCritical, "Ημερήσια κατάσταση"
        Exit Sub
    End If

    ThisWorkbook.Save
    payloadPath = Environ$("TEMP") & "\\eka_daily_input_payload_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".tsv"
    responsePath = Environ$("TEMP") & "\\eka_daily_input_response_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".json"

    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --payload " & QuoteArg(payloadPath) & _
                  " --response " & QuoteArg(responsePath)

    Application.StatusBar = "Εφαρμογή ημερήσιας κατάστασης..."
    exitCode = CreateObject("WScript.Shell").Run(commandLine, 0, True)
    Application.StatusBar = False

    If Dir$(responsePath) <> "" Then responseText = ReadUtf8Text(responsePath)
    If exitCode <> 0 Or _
       InStr(1, responseText, Chr$(34) & "ok" & Chr$(34) & ": false", vbTextCompare) > 0 Or _
       Dir$(payloadPath) = "" Then
        MsgBox "Η ημερήσια κατάσταση δεν εφαρμόστηκε." & vbCrLf & vbCrLf & _
               "Ελέγξτε τις καταχωρίσεις και δοκιμάστε ξανά." & vbCrLf & _
               "Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
               vbExclamation, "Ημερήσια κατάσταση"
        GoTo CleanUp
    End If

    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo CleanUp
    If Not RefreshOperationalViews() Then GoTo CleanUp

    replacementsNeeded = JsonLongValue(responseText, "replacement_sessions_needed")
    If replacementsNeeded > 0 Then
        If ResolveDailyReplacementsInPlace() Then
            If Not RefreshOperationalViews() Then GoTo CleanUp
            MsgBox "Η ημερήσια κατάσταση και οι αναπληρώσεις εφαρμόστηκαν επιτυχώς.", _
                   vbInformation, "Ημερήσια κατάσταση"
            GoToMaster
        End If
    Else
        MsgBox "Η ημερήσια κατάσταση εφαρμόστηκε επιτυχώς.", vbInformation, "Ημερήσια κατάσταση"
        GoToMaster
    End If

CleanUp:
    Application.StatusBar = False
    On Error Resume Next
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then Kill payloadPath
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then Kill responsePath
    On Error GoTo 0
    Exit Sub

ActionError:
    Application.StatusBar = False
    WriteDailyInputDiagnostic Err.Number, Err.Description, "daily input action", 0, "", ""
    MsgBox "Η ημερήσια κατάσταση δεν εφαρμόστηκε. Δοκιμάστε ξανά. Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.", _
           vbCritical, "Ημερήσια κατάσταση"
    Resume CleanUp
End Sub

Public Function ApplyValidatedDailyInputPayload(ByVal payloadPath As String) As Boolean
    Dim rawText As String, lines As Variant, parts As Variant
    Dim i As Long, patchCount As Long
    Dim rollbackPath As String, previousEvents As Boolean, previousScreenUpdating As Boolean
    Dim errorText As String, currentPhase As String, currentSheet As String, currentCell As String
    Dim errorNumber As Long, currentPatchIndex As Long

    On Error GoTo ApplyError

    currentPhase = "read payload"
    rawText = ReadUtf8Text(payloadPath)
    lines = Split(Replace(rawText, vbCr, ""), vbLf)
    patchCount = 0
    currentPhase = "preflight"
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo CountNext
        parts = Split(CStr(lines(i)), vbTab)
        If UBound(parts) <> 12 Or CStr(parts(0)) <> "PATCH" Then Err.Raise vbObjectError + 2301, , "Invalid DAILY_INPUT payload"
        currentPatchIndex = patchCount + 1
        currentSheet = PayloadValue(parts(1))
        currentCell = PayloadValue(parts(2))
        ValidatePatchTarget parts
        patchCount = patchCount + 1
CountNext:
    Next i
    If patchCount = 0 Then Err.Raise vbObjectError + 2302, , "Empty DAILY_INPUT payload"

    currentPhase = "create rollback"
    rollbackPath = Environ$("TEMP") & "\\eka_daily_input_rollback_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Timer * 100) & ".xlsm"
    ThisWorkbook.SaveCopyAs rollbackPath

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    currentPatchIndex = 0
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo ApplyNext
        parts = Split(CStr(lines(i)), vbTab)
        currentPatchIndex = currentPatchIndex + 1
        currentSheet = PayloadValue(parts(1))
        currentCell = PayloadValue(parts(2))
        currentPhase = "apply patch"
        ApplyPatch parts
        currentPhase = "verify patch"
        If Not VerifyPatch(parts) Then Err.Raise vbObjectError + 2303, , "DAILY_INPUT read-back mismatch"
ApplyNext:
    Next i

    currentPhase = "save"
    ThisWorkbook.Save
    ApplyValidatedDailyInputPayload = True

ApplyExit:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    On Error Resume Next
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then Kill rollbackPath
    On Error GoTo 0
    Exit Function

ApplyError:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    WriteDailyInputDiagnostic errorNumber, errorText, currentPhase, currentPatchIndex, currentSheet, currentCell
    If Len(rollbackPath) > 0 And Dir$(rollbackPath) <> "" Then
        RestorePatchesFromRollback rollbackPath, lines
        ThisWorkbook.Save
    End If
    On Error GoTo 0
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    MsgBox "Η ημερήσια κατάσταση δεν εφαρμόστηκε και οι αλλαγές αναιρέθηκαν." & vbCrLf & _
           "Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.", vbCritical, "Ημερήσια κατάσταση"
    Resume ApplyExit
End Function

Private Function RefreshOperationalViews() As Boolean
    Dim master As Worksheet, daily As Worksheet, refreshShape As Shape
    On Error GoTo RefreshError

    Set master = ThisWorkbook.Worksheets("MASTER_SCHEDULE")
    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    RenderTherapistWorkloadStrip
    Set refreshShape = master.Shapes("ekaWorkload_Refresh")
    ApplyMasterView
    RefreshOperationalViews = True
    Exit Function

RefreshError:
    WriteDailyInputDiagnostic Err.Number, Err.Description, "refresh operational views", 0, "MASTER_SCHEDULE", ""
    MsgBox "Οι αλλαγές αποθηκεύτηκαν, αλλά η κεντρική εικόνα δεν ενημερώθηκε σωστά." & vbCrLf & _
           "Επικοινωνήστε με τον διαχειριστή πριν συνεχίσετε.", vbCritical, "Ημερήσια κατάσταση"
End Function

Private Sub ValidatePatchTarget(ByVal parts As Variant)
    Dim ws As Worksheet, cellAddress As String
    Dim target As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PayloadValue(parts(1)))
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise vbObjectError + 2310, , "Missing target sheet"
    cellAddress = PayloadValue(parts(2))
    If Len(cellAddress) = 0 Then Err.Raise vbObjectError + 2311, , "Missing target cell"
    On Error GoTo InvalidCell
    Set target = ws.Range(cellAddress)
    Exit Sub
InvalidCell:
    Err.Raise vbObjectError + 2312, , "Invalid target cell"
End Sub

Private Sub ApplyPatch(ByVal parts As Variant)
    Dim ws As Worksheet, cell As Range
    Dim intent As String, valueText As String
    Dim wrapToken As String, minFontToken As String, textRuns As String
    Dim fontSizeValue As Variant

    Set ws = ThisWorkbook.Worksheets(PayloadValue(parts(1)))
    Set cell = ws.Range(PayloadValue(parts(2)))
    intent = LCase$(PayloadValue(parts(3)))
    valueText = PayloadValue(parts(4))

    Select Case intent
        Case "clear": cell.ClearContents
        Case "value": cell.Value = valueText
        Case "presentation"
        Case Else: Err.Raise vbObjectError + 2320, , "Unsupported write intent"
    End Select

    wrapToken = PayloadValue(parts(9))
    If wrapToken = "1" Then cell.WrapText = True
    If wrapToken = "0" Then cell.WrapText = False

    minFontToken = PayloadValue(parts(10))
    If Len(minFontToken) > 0 Then
        fontSizeValue = cell.Font.Size
        If IsNull(fontSizeValue) Then
            cell.Font.Size = CDbl(minFontToken)
        ElseIf CDbl(fontSizeValue) < CDbl(minFontToken) Then
            cell.Font.Size = CDbl(minFontToken)
        End If
    End If

    If PayloadValue(parts(5)) = "1" Then cell.Font.Strikethrough = True
    ApplyWholeFontRole cell, PayloadValue(parts(6))
    ApplyFillRole cell, PayloadValue(parts(7))
    ApplyBorderRole cell, PayloadValue(parts(8))
    textRuns = PayloadValue(parts(12))
    If Len(textRuns) > 0 Then ApplyTextRuns cell, textRuns, PayloadValue(parts(6))
End Sub

Private Function VerifyPatch(ByVal parts As Variant) As Boolean
    Dim ws As Worksheet, cell As Range
    Dim intent As String, expected As String, wrapToken As String, minFontToken As String
    Dim textRuns As String, fontSizeValue As Variant

    Set ws = ThisWorkbook.Worksheets(PayloadValue(parts(1)))
    Set cell = ws.Range(PayloadValue(parts(2)))
    intent = LCase$(PayloadValue(parts(3)))
    expected = PayloadValue(parts(4))
    textRuns = PayloadValue(parts(12))

    If intent = "value" Then
        If CStr(cell.Value) <> expected Then Exit Function
    ElseIf intent = "clear" Then
        If Len(CStr(cell.Value)) <> 0 Then Exit Function
    End If

    wrapToken = PayloadValue(parts(9))
    If wrapToken = "1" And cell.WrapText <> True Then Exit Function
    If wrapToken = "0" And cell.WrapText <> False Then Exit Function

    minFontToken = PayloadValue(parts(10))
    If Len(minFontToken) > 0 And Len(textRuns) = 0 Then
        fontSizeValue = cell.Font.Size
        If IsNull(fontSizeValue) Then Exit Function
        If CDbl(fontSizeValue) < CDbl(minFontToken) Then Exit Function
    End If

    If Not VerifyFillRole(cell, PayloadValue(parts(7))) Then Exit Function
    If Not VerifyWholeFontRole(cell, PayloadValue(parts(6))) Then Exit Function
    If Not VerifyBorderRole(cell, PayloadValue(parts(8))) Then Exit Function
    If Len(textRuns) > 0 Then
        If Not VerifyTextRuns(cell, textRuns) Then Exit Function
    End If
    VerifyPatch = True
End Function

Private Sub ApplyWholeFontRole(ByVal cell As Range, ByVal role As String)
    Select Case LCase$(role)
        Case "student_active_green", "outpatient_dark_green": cell.Font.Color = RGB(0, 128, 0)
        Case "robotic_orange": cell.Font.Color = RGB(255, 140, 0)
        Case "muted": cell.Font.Color = RGB(102, 102, 102)
        Case "default": cell.Font.Color = RGB(0, 0, 0)
    End Select
End Sub

Private Function VerifyWholeFontRole(ByVal cell As Range, ByVal role As String) As Boolean
    Select Case LCase$(role)
        Case "student_active_green", "outpatient_dark_green": VerifyWholeFontRole = (CLng(cell.Font.Color) = CLng(RGB(0, 128, 0)))
        Case "robotic_orange": VerifyWholeFontRole = (CLng(cell.Font.Color) = CLng(RGB(255, 140, 0)))
        Case "muted": VerifyWholeFontRole = (CLng(cell.Font.Color) = CLng(RGB(102, 102, 102)))
        Case "default": VerifyWholeFontRole = (CLng(cell.Font.Color) = CLng(RGB(0, 0, 0)))
        Case Else: VerifyWholeFontRole = True
    End Select
End Function

Private Sub ApplyFillRole(ByVal cell As Range, ByVal role As String)
    Select Case LCase$(role)
        Case "infectious_yellow": cell.Interior.Color = RGB(255, 255, 67)
        Case "robotic_pink": cell.Interior.Color = RGB(248, 203, 173)
        Case "outpatient_light_blue": cell.Interior.Color = RGB(157, 181, 143)
        Case "clear_fill": cell.Interior.Pattern = xlPatternNone
    End Select
End Sub

Private Function VerifyFillRole(ByVal cell As Range, ByVal role As String) As Boolean
    Select Case LCase$(role)
        Case "infectious_yellow": VerifyFillRole = (CLng(cell.Interior.Color) = CLng(RGB(255, 255, 67)))
        Case "robotic_pink": VerifyFillRole = (CLng(cell.Interior.Color) = CLng(RGB(248, 203, 173)))
        Case "outpatient_light_blue": VerifyFillRole = (CLng(cell.Interior.Color) = CLng(RGB(157, 181, 143)))
        Case "clear_fill": VerifyFillRole = (cell.Interior.Pattern = xlPatternNone)
        Case Else: VerifyFillRole = True
    End Select
End Function

Private Sub ApplyBorderRole(ByVal cell As Range, ByVal role As String)
    Dim edge As Variant
    If LCase$(role) <> "infectious_yellow" Then Exit Sub
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With cell.Borders(CLng(edge))
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(255, 255, 67)
        End With
    Next edge
End Sub

Private Function VerifyBorderRole(ByVal cell As Range, ByVal role As String) As Boolean
    If LCase$(role) <> "infectious_yellow" Then VerifyBorderRole = True: Exit Function
    VerifyBorderRole = (CLng(cell.Borders(xlEdgeLeft).Color) = CLng(RGB(255, 255, 67)))
End Function

Private Sub ApplyTextRuns(ByVal cell As Range, ByVal token As String, ByVal wholeFontRole As String)
    Dim runs As Variant, fields As Variant, item As Variant
    Dim startPos As Long, runLength As Long, role As String

    cell.Font.Strikethrough = False
    cell.Font.Italic = False
    If Len(wholeFontRole) = 0 Then cell.Font.Color = RGB(0, 0, 0)

    runs = Split(token, ";")
    For Each item In runs
        fields = Split(CStr(item), ":")
        If UBound(fields) <> 4 Then Err.Raise vbObjectError + 2330, , "Invalid rich-text run"
        startPos = CLng(fields(0))
        runLength = CLng(fields(1))
        With cell.Characters(startPos, runLength).Font
            .Strikethrough = (fields(2) = "1")
            .Italic = (fields(3) = "1")
            role = CStr(fields(4))
            Select Case LCase$(role)
                Case "student_active_green", "outpatient_dark_green": .Color = RGB(0, 128, 0)
                Case "robotic_orange": .Color = RGB(255, 140, 0)
                Case "muted": .Color = RGB(102, 102, 102)
                Case "default": .Color = RGB(0, 0, 0)
            End Select
        End With
    Next item
End Sub

Private Function VerifyTextRuns(ByVal cell As Range, ByVal token As String) As Boolean
    Dim runs As Variant, fields As Variant, item As Variant
    Dim startPos As Long, runLength As Long, role As String
    Dim chars As Characters

    runs = Split(token, ";")
    For Each item In runs
        fields = Split(CStr(item), ":")
        If UBound(fields) <> 4 Then Exit Function
        startPos = CLng(fields(0))
        runLength = CLng(fields(1))
        Set chars = cell.Characters(startPos, runLength)
        If CBool(chars.Font.Strikethrough) <> (fields(2) = "1") Then Exit Function
        If CBool(chars.Font.Italic) <> (fields(3) = "1") Then Exit Function
        role = LCase$(CStr(fields(4)))
        Select Case role
            Case "student_active_green", "outpatient_dark_green"
                If CLng(chars.Font.Color) <> CLng(RGB(0, 128, 0)) Then Exit Function
            Case "robotic_orange"
                If CLng(chars.Font.Color) <> CLng(RGB(255, 140, 0)) Then Exit Function
            Case "muted"
                If CLng(chars.Font.Color) <> CLng(RGB(102, 102, 102)) Then Exit Function
            Case "default"
                If CLng(chars.Font.Color) <> CLng(RGB(0, 0, 0)) Then Exit Function
        End Select
    Next item
    VerifyTextRuns = True
End Function

Private Sub RestorePatchesFromRollback(ByVal rollbackPath As String, ByVal lines As Variant)
    Dim rollbackBook As Workbook, parts As Variant
    Dim i As Long, sheetName As String, cellAddress As String
    Dim previousAlerts As Boolean

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set rollbackBook = Application.Workbooks.Open(Filename:=rollbackPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error Resume Next
    If rollbackBook.Windows.Count > 0 Then rollbackBook.Windows(1).Visible = False
    On Error GoTo 0

    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo RestoreNext
        parts = Split(CStr(lines(i)), vbTab)
        sheetName = PayloadValue(parts(1))
        cellAddress = PayloadValue(parts(2))
        rollbackBook.Worksheets(sheetName).Range(cellAddress).Copy Destination:=ThisWorkbook.Worksheets(sheetName).Range(cellAddress)
RestoreNext:
    Next i
    Application.CutCopyMode = False
    rollbackBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
End Sub

Private Sub WriteDailyInputDiagnostic( _
    ByVal errorNumber As Long, _
    ByVal errorText As String, _
    ByVal phase As String, _
    ByVal patchIndex As Long, _
    ByVal sheetName As String, _
    ByVal cellAddress As String)
    Dim filePath As String, handle As Integer

    filePath = Environ$("TEMP") & "\\eka_daily_input_last_error.log"
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #handle, "phase=" & phase
    Print #handle, "patch_index=" & CStr(patchIndex)
    Print #handle, "sheet=" & sheetName
    Print #handle, "cell=" & cellAddress
    Print #handle, "error_number=" & CStr(errorNumber)
    Print #handle, "error=" & errorText
    Close #handle
End Sub

Private Function PayloadValue(ByVal raw As String) As String
    Dim i As Long, ch As String, result As String, escaped As Boolean
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If escaped Then
            Select Case ch
                Case "t": result = result & vbTab
                Case "r": result = result & vbCr
                Case "n": result = result & vbLf
                Case "\\": result = result & "\\"
                Case Else: result = result & ch
            End Select
            escaped = False
        ElseIf ch = "\\" Then
            escaped = True
        Else
            result = result & ch
        End If
    Next i
    If escaped Then result = result & "\\"
    PayloadValue = result
End Function

Private Function JsonLongValue(ByVal jsonText As String, ByVal key As String) As Long
    Dim marker As String, startPos As Long, endPos As Long, rawValue As String
    marker = Chr$(34) & key & Chr$(34) & ":"
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(marker)
    Do While startPos <= Len(jsonText) And Mid$(jsonText, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    endPos = startPos
    Do While endPos <= Len(jsonText) And Mid$(jsonText, endPos, 1) Like "[0-9]"
        endPos = endPos + 1
    Loop
    rawValue = Mid$(jsonText, startPos, endPos - startPos)
    If Len(rawValue) > 0 Then JsonLongValue = CLng(rawValue)
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

Private Function ResolveDailyInputPreparationScript() As String
    Dim candidate As String
    candidate = ThisWorkbook.Path & "\\Python\\tools\\prepare_daily_input_in_place.py"
    If Dir$(candidate) <> "" Then ResolveDailyInputPreparationScript = candidate: Exit Function
    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\prepare_daily_input_in_place.py"
    If Dir$(candidate) <> "" Then ResolveDailyInputPreparationScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
'''

# VBA string literals do not use backslash escaping. Normalize the Python raw-string
# source before installation so paths and one-character payload escape parsing use
# a single Windows separator/backslash at runtime.
DAILY_INPUT_MODULE_CODE = DAILY_INPUT_MODULE_CODE.replace("\\\\", "\\")


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def install_daily_input_action(vbproject) -> None:
    _remove_component_if_present(vbproject, DAILY_INPUT_MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = DAILY_INPUT_MODULE_NAME
    module.CodeModule.AddFromString(DAILY_INPUT_MODULE_CODE)

    install_global_replacement_action(vbproject)
    install_double_booking_review(vbproject)

    _remove_component_if_present(vbproject, WORKLOAD_MODULE_NAME)
    workload_module = vbproject.VBComponents.Add(1)
    workload_module.Name = WORKLOAD_MODULE_NAME
    workload_module.CodeModule.AddFromString(WORKLOAD_MODULE_CODE)
