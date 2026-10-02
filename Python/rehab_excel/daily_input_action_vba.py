from __future__ import annotations

DAILY_INPUT_MODULE_NAME = "modDailyInputAction"
DAILY_INPUT_MACRO_NAME = "ApplyDailyInputPreview"

DAILY_INPUT_MODULE_CODE = r'''Option Explicit

Public Sub ApplyDailyInputPreview()
    Dim scriptPath As String
    Dim outputPath As String
    Dim commandLine As String
    Dim shell As Object
    Dim exec As Object
    Dim stdoutText As String
    Dim stderrText As String

    On Error GoTo ActionError

    If ThisWorkbook.ReadOnly Then
        MsgBox "Το αρχείο είναι μόνο για ανάγνωση. Δεν μπορεί να αποθηκευτεί το DAILY_INPUT πριν από την εφαρμογή.", vbExclamation, "DAILY_INPUT"
        Exit Sub
    End If

    scriptPath = ResolveDailyInputScript()
    If Len(scriptPath) = 0 Then
        MsgBox "Δεν βρέθηκε το εργαλείο εφαρμογής DAILY_INPUT.", vbCritical, "DAILY_INPUT"
        Exit Sub
    End If

    ThisWorkbook.Save

    outputPath = ThisWorkbook.Path & "\\DAILY_INPUT_APPLIED_PREVIEW.xlsm"
    commandLine = QuoteArg("python") & " " & QuoteArg(scriptPath) & _
                  " --input " & QuoteArg(ThisWorkbook.FullName) & _
                  " --output " & QuoteArg(outputPath) & _
                  " --overwrite"

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(commandLine)

    Do While exec.Status = 0
        DoEvents
    Loop

    stdoutText = exec.StdOut.ReadAll
    stderrText = exec.StdErr.ReadAll

    If exec.ExitCode <> 0 Then
        If Len(Trim$(stderrText)) > 0 Then
            MsgBox "Η εφαρμογή του DAILY_INPUT δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & stderrText, vbExclamation, "DAILY_INPUT"
        Else
            MsgBox "Η εφαρμογή του DAILY_INPUT δεν ολοκληρώθηκε:" & vbCrLf & vbCrLf & stdoutText, vbExclamation, "DAILY_INPUT"
        End If
        Exit Sub
    End If

    MsgBox "Δημιουργήθηκε ασφαλές preview της ημερήσιας κατάστασης." & vbCrLf & _
           "Αρχείο: " & outputPath, vbInformation, "DAILY_INPUT"
    Exit Sub

ActionError:
    MsgBox "Δεν ήταν δυνατή η εφαρμογή του DAILY_INPUT: " & Err.Description, vbCritical, "DAILY_INPUT"
End Sub

Private Function ResolveDailyInputScript() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\\Python\\tools\\apply_daily_input.py"
    If Dir$(candidate) <> "" Then
        ResolveDailyInputScript = candidate
        Exit Function
    End If

    candidate = ThisWorkbook.Path & "\\..\\..\\Python\\tools\\apply_daily_input.py"
    If Dir$(candidate) <> "" Then
        ResolveDailyInputScript = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(candidate)
    End If
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
'''


def install_daily_input_action(vbproject) -> None:
    try:
        existing = vbproject.VBComponents(DAILY_INPUT_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        vbproject.VBComponents.Remove(existing)

    module = vbproject.VBComponents.Add(1)
    module.Name = DAILY_INPUT_MODULE_NAME
    module.CodeModule.AddFromString(DAILY_INPUT_MODULE_CODE)
