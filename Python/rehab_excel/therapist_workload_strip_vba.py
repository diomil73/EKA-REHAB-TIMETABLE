from __future__ import annotations

MODULE_NAME = "modTherapistWorkloadStrip"
SHAPE_PREFIX = "ekaWorkload_"
BUTTON_NAME = SHAPE_PREFIX + "Refresh"

MODULE_CODE = r'''Option Explicit

Private Const DEFAULT_CAPACITY As Long = 6

Public Function TherapistCurrentWorkload(ByVal providerName As String) As Long
    Dim daily As Worksheet
    Dim headerRows As Variant
    Dim firstRows As Variant
    Dim lastRows As Variant
    Dim blockIndex As Long
    Dim col As Long
    Dim rowIndex As Long
    Dim providerKey As String
    Dim headerKey As String
    Dim timeValue As String
    Dim providerTimes As Object

    Set daily = ThisWorkbook.Worksheets("THERAPIST_DAILY")
    providerKey = NormalizeProviderName(providerName)
    If Len(providerKey) = 0 Then Exit Function

    headerRows = Array(1, 11)
    firstRows = Array(2, 12)
    lastRows = Array(8, 18)
    Set providerTimes = CreateObject("Scripting.Dictionary")

    For blockIndex = LBound(headerRows) To UBound(headerRows)
        For col = 2 To 10
            headerKey = NormalizeProviderName(CStr(daily.Cells(headerRows(blockIndex), col).Value))
            If headerKey = providerKey Then
                For rowIndex = firstRows(blockIndex) To lastRows(blockIndex)
                    If Len(Trim$(CStr(daily.Cells(rowIndex, col).Value))) > 0 Then
                        timeValue = TimeKey(daily.Cells(rowIndex, 1).Value)
                        If Len(timeValue) = 0 Then timeValue = TimeKey(daily.Cells(rowIndex, 1).Text)
                        If Len(timeValue) > 0 Then providerTimes(timeValue) = True
                    End If
                Next rowIndex
            End If
        Next col
    Next blockIndex

    TherapistCurrentWorkload = providerTimes.Count
End Function

Public Function TherapistWorkloadCapacity() As Long
    TherapistWorkloadCapacity = DEFAULT_CAPACITY
End Function

Private Function NormalizeProviderName(ByVal valueText As String) As String
    Dim text As String
    text = LCase$(Trim$(valueText))
    text = Replace(text, "ά", "α")
    text = Replace(text, "έ", "ε")
    text = Replace(text, "ή", "η")
    text = Replace(text, "ί", "ι")
    text = Replace(text, "ϊ", "ι")
    text = Replace(text, "ΐ", "ι")
    text = Replace(text, "ό", "ο")
    text = Replace(text, "ύ", "υ")
    text = Replace(text, "ϋ", "υ")
    text = Replace(text, "ΰ", "υ")
    text = Replace(text, "ώ", "ω")
    text = Replace(text, "ς", "σ")
    text = Replace(text, ".", "")
    text = Replace(text, " ", "")
    NormalizeProviderName = text
End Function

Private Function TimeKey(ByVal rawValue As Variant) As String
    Dim text As String
    On Error GoTo InvalidTime

    If IsDate(rawValue) Then
        TimeKey = Format$(CDate(rawValue), "hh:mm")
        Exit Function
    End If
    If IsNumeric(rawValue) Then
        TimeKey = Format$(CDate(CDbl(rawValue)), "hh:mm")
        Exit Function
    End If

    text = Replace(Trim$(CStr(rawValue)), ".", ":")
    If IsDate(text) Then TimeKey = Format$(TimeValue(text), "hh:mm")
    Exit Function

InvalidTime:
    TimeKey = ""
End Function
'''


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _delete_legacy_master_workload_shapes(workbook) -> None:
    master = workbook.Worksheets("MASTER_SCHEDULE")
    names: list[str] = []
    for index in range(1, int(master.Shapes.Count) + 1):
        name = str(master.Shapes(index).Name)
        if name.startswith(SHAPE_PREFIX):
            names.append(name)
    for name in names:
        master.Shapes(name).Delete()


def install_therapist_workload_strip(vbproject, workbook) -> None:
    """Install workload calculation as a backend dependency, not MASTER UI.

    Authoritative UI direction places current workload and workload history in
    the therapist's personal card. MASTER must not publish workload badges or a
    workload button. Legacy recovery shapes are removed during installation.
    """

    _remove_component_if_present(vbproject, MODULE_NAME)
    module = vbproject.VBComponents.Add(1)
    module.Name = MODULE_NAME
    module.CodeModule.AddFromString(MODULE_CODE)
    _delete_legacy_master_workload_shapes(workbook)
