from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]


class CloseDayDiagnosticInstallError(RuntimeError):
    pass


MODULE_NAME = "modCloseDayAction"


def _module_text(code_module) -> str:
    count = int(code_module.CountOfLines)
    if count <= 0:
        return ""
    return str(code_module.Lines(1, count))


def _replace_once(text: str, old: str, new: str, *, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise CloseDayDiagnosticInstallError(
            f"Expected exactly one {label} target, found {count}"
        )
    return text.replace(old, new, 1)


def _assert_markers(source: str) -> None:
    markers = (
        'persistStage = "patient-write-readback row=" & CStr(i)',
        '" col=" & CStr(mismatchCol + 1)',
        '" expected=[" & expectedValue & "] actual=[" & actualValueText & "] type="',
        '"Diagnostic stage: " & persistStage',
        "errorNumber = Err.Number",
        "errorDescription = Err.Description",
    )
    for marker in markers:
        if marker not in source:
            raise CloseDayDiagnosticInstallError(
                f"Diagnostic marker not confirmed: {marker}"
            )


def install_close_day_persistence_diagnostic(workbook_path: str | Path) -> dict[str, object]:
    if sys.platform != "win32":
        raise CloseDayDiagnosticInstallError(
            "Close-day diagnostic installation requires Windows + Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise CloseDayDiagnosticInstallError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise CloseDayDiagnosticInstallError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise CloseDayDiagnosticInstallError("Workbook must be an .xlsm file")

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)

        try:
            component = workbook.VBProject.VBComponents(MODULE_NAME)
        except Exception as exc:
            raise CloseDayDiagnosticInstallError(
                f"VBA module not found: {MODULE_NAME}"
            ) from exc

        code_module = component.CodeModule
        source = _module_text(code_module)

        source = _replace_once(
            source,
            "    Dim wsSnapshot As Worksheet, wsTherapist As Worksheet, wsPatient As Worksheet, wsEvent As Worksheet\r\n",
            "    Dim wsSnapshot As Worksheet, wsTherapist As Worksheet, wsPatient As Worksheet, wsEvent As Worksheet\r\n"
            "    Dim persistStage As String, errorDescription As String\r\n"
            "    Dim errorNumber As Long, mismatchCol As Long\r\n"
            "    Dim expectedValue As String, actualValueText As String\r\n",
            label="diagnostic declaration",
        )

        replacements = (
            ("    rawText = ReadUtf8Text(payloadPath)\r\n", "    persistStage = \"read-payload\"\r\n    rawText = ReadUtf8Text(payloadPath)\r\n"),
            ("    If Not PreflightSnapshot(snapshotParts) Then Exit Function\r\n", "    persistStage = \"preflight\"\r\n    If Not PreflightSnapshot(snapshotParts) Then Exit Function\r\n"),
            ("    Set wsSnapshot = EnsureHistorySheet(SNAPSHOT_SHEET, SnapshotHeaders(), snapshotCreated)\r\n", "    persistStage = \"ensure-history-sheets\"\r\n    Set wsSnapshot = EnsureHistorySheet(SNAPSHOT_SHEET, SnapshotHeaders(), snapshotCreated)\r\n"),
            ("    WritePayloadFields wsSnapshot, snapshotRow, snapshotParts, 1, 15\r\n", "    persistStage = \"snapshot-write-readback\"\r\n    WritePayloadFields wsSnapshot, snapshotRow, snapshotParts, 1, 15\r\n"),
            ("    For i = 1 To therapistCount\r\n        parts = Split(therapistLines(i), vbTab)\r\n", "    persistStage = \"therapist-write-readback\"\r\n    For i = 1 To therapistCount\r\n        parts = Split(therapistLines(i), vbTab)\r\n"),
            ("    For i = 1 To patientCount\r\n        parts = Split(patientLines(i), vbTab)\r\n", "    persistStage = \"patient-write-readback\"\r\n    For i = 1 To patientCount\r\n        persistStage = \"patient-write-readback row=\" & CStr(i)\r\n        parts = Split(patientLines(i), vbTab)\r\n"),
            ("    WritePayloadFields wsEvent, eventRow, eventParts, 1, 12\r\n", "    persistStage = \"event-write-readback\"\r\n    WritePayloadFields wsEvent, eventRow, eventParts, 1, 12\r\n"),
            ("    FinalizeTechnicalSheet wsSnapshot\r\n", "    persistStage = \"finalize-protect-save\"\r\n    FinalizeTechnicalSheet wsSnapshot\r\n"),
        )
        for old, new in replacements:
            source = _replace_once(source, old, new, label=old.strip())

        old_patient_check = (
            "        If Not RowMatchesPayload(wsPatient, patientStartRow + i - 1, parts, 1, 14) Then Err.Raise vbObjectError + 2113, , \"Patient history read-back mismatch\"\r\n"
        )
        new_patient_check = (
            "        If Not RowMatchesPayload(wsPatient, patientStartRow + i - 1, parts, 1, 14) Then\r\n"
            "            For mismatchCol = 0 To 13\r\n"
            "                expectedValue = PayloadValue(parts(1 + mismatchCol))\r\n"
            "                If Not CellMatchesPayload(wsPatient.Cells(patientStartRow + i - 1, mismatchCol + 1).Value, expectedValue) Then\r\n"
            "                    actualValueText = CStr(wsPatient.Cells(patientStartRow + i - 1, mismatchCol + 1).Value)\r\n"
            "                    Err.Raise vbObjectError + 2113, , \"Patient history read-back mismatch row=\" & CStr(i) & _\r\n"
            "                        \" col=\" & CStr(mismatchCol + 1) & _\r\n"
            "                        \" expected=[\" & expectedValue & \"] actual=[\" & actualValueText & \"] type=\" & _\r\n"
            "                        TypeName(wsPatient.Cells(patientStartRow + i - 1, mismatchCol + 1).Value)\r\n"
            "                End If\r\n"
            "            Next mismatchCol\r\n"
            "            Err.Raise vbObjectError + 2113, , \"Patient history read-back mismatch with no differing field identified\"\r\n"
            "        End If\r\n"
        )
        source = _replace_once(
            source,
            old_patient_check,
            new_patient_check,
            label="patient mismatch detail",
        )

        old_handler = (
            "PersistError:\r\n"
            "    On Error Resume Next\r\n"
        )
        new_handler = (
            "PersistError:\r\n"
            "    errorNumber = Err.Number\r\n"
            "    errorDescription = Err.Description\r\n"
            "    On Error Resume Next\r\n"
        )
        source = _replace_once(source, old_handler, new_handler, label="PersistError capture")

        old_message = (
            "    MsgBox \"Η ημέρα δεν έκλεισε. Δεν αποθηκεύτηκαν αλλαγές. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.\", vbCritical, \"Κλείσιμο ημέρας\"\r\n"
        )
        new_message = (
            "    MsgBox \"Η ημέρα δεν έκλεισε. Δεν αποθηκεύτηκαν αλλαγές.\" & vbCrLf & vbCrLf & _\r\n"
            "           \"Diagnostic stage: \" & persistStage & vbCrLf & _\r\n"
            "           \"VBA error: \" & CStr(errorNumber) & vbCrLf & _\r\n"
            "           \"Description: \" & errorDescription, vbCritical, \"Κλείσιμο ημέρας\"\r\n"
        )
        source = _replace_once(source, old_message, new_message, label="diagnostic failure message")

        code_module.DeleteLines(1, int(code_module.CountOfLines))
        code_module.AddFromString(source)
        _assert_markers(_module_text(code_module))
        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None

        # Reopen from disk and verify that the diagnostic VBA actually persisted.
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        try:
            persisted_component = workbook.VBProject.VBComponents(MODULE_NAME)
        except Exception as exc:
            raise CloseDayDiagnosticInstallError(
                f"Persisted VBA module not found after reopen: {MODULE_NAME}"
            ) from exc
        persisted_source = _module_text(persisted_component.CodeModule)
        _assert_markers(persisted_source)

        return {
            "ok": True,
            "workbook": str(path),
            "module": MODULE_NAME,
            "persisted_reopen_verified": True,
        }
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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Inject stage/error diagnostics into close-day VBA on a disposable XLSM smoke workbook."
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        result = install_close_day_persistence_diagnostic(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("CLOSE DAY PERSISTENCE DIAGNOSTIC INSTALL OK")
    print(f"Workbook: {result['workbook']}")
    print(f"Module: {result['module']}")
    print("Diagnostic stages: installed")
    print("Patient mismatch detail: installed")
    print("Persisted VBA reopen verification: OK")
    print("Rollback semantics: unchanged")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
