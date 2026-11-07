from __future__ import annotations

"""Safely patch the six-step refresh UI in a fresh, disposable native workbook.

Only known successful MsgBox sites are suppressed during RefreshAll. All
standalone procedure messages and all error/confirmation dialogs remain intact.
The installer fails closed on any unexpected VBA source contract.
"""

import argparse
from pathlib import Path
import re
import sys

EXPECTED_REFRESH = (
    "Public Sub RefreshAll()",
    "Application.ScreenUpdating = False",
    "Application.EnableEvents = False",
    "Call RefreshTherapistDaily",
    "Call RefreshTherapistAttendanceAndConflicts",
    "Call BuildReplacementSuggestions_V4",
    "Call ColorInfectiousInMaster",
    "Call ColorReplacementMarks",
    "Call RestoreReturnedTherapists",
    "Application.EnableEvents = True",
    "Application.ScreenUpdating = True",
    'MsgBox "Όλα ενημερώθηκαν.", vbInformation',
    "End Sub",
)
SUCCESS_SITES = {
    "Module1": ('MsgBox "Το THERAPIST_DAILY ανανεώθηκε.", vbInformation',),
    "Module2": ('MsgBox "THERAPIST_ATTENDANCE και CONFLICT_LOG ανανεώθηκαν.", vbInformation',),
    "Module4": (
        'MsgBox "Οι προτάσεις αναπλήρωσης v4 δημιουργήθηκαν.", vbInformation',
        'MsgBox "Κουμπιά που δημιουργήθηκαν: " & createdCount, vbInformation',
    ),
    "Module10": ('MsgBox "Ολοκληρώθηκε ο έλεγχος επαναφοράς.", vbInformation',),
}
FLAG = "EKA_SilentRefreshSuccess"
REFRESH_BLOCK = re.compile(
    r"(?ms)^Public Sub RefreshAll\(\)\r?\n.*?^End Sub\s*$",
)


class RefreshInstallError(RuntimeError):
    pass


def _replace_unique(source: str, target: str, replacement: str, module: str) -> str:
    pattern = re.compile(r"(?m)^(?P<indent>[ \t]*)" + re.escape(target) + r"[ \t]*$")
    matches = list(pattern.finditer(source))
    if len(matches) != 1:
        raise RefreshInstallError(
            f"{module}: expected exactly one matching successful notification {target!r}; found {len(matches)}"
        )
    return pattern.sub(lambda m: m.group("indent") + replacement, source, count=1)


def plan_patch(modules: dict[str, str]) -> dict[str, str]:
    """Pure all-or-nothing validation before any Excel COM edits."""
    required = {"Module3", *SUCCESS_SITES}
    if set(modules) != required:
        raise RefreshInstallError(f"Unexpected VBA source selection: {sorted(modules)}")
    source = modules["Module3"].replace("\r\n", "\n")
    if FLAG.casefold() in source.casefold():
        raise RefreshInstallError("Refresh success flag already exists; refusing reinjection")
    blocks = list(REFRESH_BLOCK.finditer(source))
    if len(blocks) != 1:
        raise RefreshInstallError("RefreshAll procedure is not uniquely located")
    block = blocks[0]
    actual = tuple(line.strip() for line in block.group().splitlines() if line.strip())
    if actual != EXPECTED_REFRESH:
        raise RefreshInstallError(
            "RefreshAll VBA body differs from the audited six-step procedure; refusing patch"
        )
    revised = """Public Sub RefreshAll()
    Dim priorEvents As Boolean, priorScreenUpdating As Boolean
    Dim failureNumber As Long, failureText As String
    If EKA_SilentRefreshSuccess Then
        MsgBox "Η ανανέωση βρίσκεται ήδη σε εξέλιξη.", vbExclamation
        Exit Sub
    End If
    priorEvents = Application.EnableEvents
    priorScreenUpdating = Application.ScreenUpdating
    On Error GoTo RefreshFailure
    EKA_SilentRefreshSuccess = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call RefreshTherapistDaily
    Call RefreshTherapistAttendanceAndConflicts
    Call BuildReplacementSuggestions_V4
    Call ColorInfectiousInMaster
    Call ColorReplacementMarks
    Call RestoreReturnedTherapists
    EKA_SilentRefreshSuccess = False
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorScreenUpdating
    MsgBox "Όλα ενημερώθηκαν.", vbInformation
    Exit Sub
RefreshFailure:
    failureNumber = Err.Number
    failureText = Err.Description
    EKA_SilentRefreshSuccess = False
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorScreenUpdating
    MsgBox "Η ανανέωση δεν ολοκληρώθηκε: " & CStr(failureNumber) & _
        " / " & failureText, vbCritical, "Ανανέωση"
End Sub"""
    source = source[:block.start()] + revised + source[block.end():]
    source = "Public EKA_SilentRefreshSuccess As Boolean\n" + source
    output = {"Module3": source}
    for module, notifications in SUCCESS_SITES.items():
        updated = modules[module].replace("\r\n", "\n")
        if FLAG.casefold() in updated.casefold():
            raise RefreshInstallError(f"{module} already has a refresh suppression flag")
        for target in notifications:
            updated = _replace_unique(
                updated, target, f"If Not {FLAG} Then {target}", module
            )
        output[module] = updated
    return output


def _read_modules(workbook) -> dict[str, str]:
    return {
        name: str(workbook.VBProject.VBComponents(name).CodeModule.Lines(
            1, workbook.VBProject.VBComponents(name).CodeModule.CountOfLines
        ))
        for name in ("Module3", *SUCCESS_SITES)
    }


def _audit(workbook, expected: dict[str, str]) -> None:
    if not bool(workbook.ReadOnly):
        raise RefreshInstallError("Persisted verification must be read-only")
    for module, wanted in expected.items():
        actual = str(workbook.VBProject.VBComponents(module).CodeModule.Lines(
            1, workbook.VBProject.VBComponents(module).CodeModule.CountOfLines
        )).replace("\r\n", "\n")
        if actual.strip() != wanted.strip():
            raise RefreshInstallError(f"Persisted VBA mismatch in {module}")
    source = expected["Module3"]
    if "Call RestoreReturnedTherapists" not in source or "RefreshFailure:" not in source:
        raise RefreshInstallError("Refresh pipeline / error guard missing")


def install(path: Path) -> None:
    if sys.platform != "win32" or not path.is_file() or path.suffix.casefold() != ".xlsm":
        raise RefreshInstallError("Existing fresh XLSM on Windows required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RefreshInstallError("pywin32 required") from exc
    excel = win32com.client.DispatchEx("Excel.Application")
    workbook = None
    patches: dict[str, str] = {}
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        if bool(workbook.ReadOnly):
            raise RefreshInstallError("Workbook unexpectedly read-only")
        patches = plan_patch(_read_modules(workbook))
        for module, text in patches.items():
            code = workbook.VBProject.VBComponents(module).CodeModule
            code.DeleteLines(1, code.CountOfLines)
            code.AddFromString(text)
        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()

    excel = win32com.client.DispatchEx("Excel.Application")
    workbook = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        _audit(workbook, patches)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(description="Install H8 single-summary RefreshAll VBA in disposable XLSM")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    try:
        install(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        print("Freeze failed destination. Do not retry installer on same XLSM.")
        return 2
    print("H8 REFRESH ONE-MESSAGE INSTALL OK")
    print("Delegated refresh success dialogs suppressed only during RefreshAll")
    print("Standalone notifications retained; RefreshAll errors retain one final error dialog")
    print("Persisted VBA reopen verification: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
