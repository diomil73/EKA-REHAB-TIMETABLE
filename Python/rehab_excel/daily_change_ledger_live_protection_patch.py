from __future__ import annotations

from pathlib import Path
import sys


MODULE_NAME = "modDailyChangeLedger"


class DailyChangeLedgerLiveProtectionPatchError(RuntimeError):
    pass


def patch_ledger_module_source(source: str) -> str:
    """Make the THERAPIST_DAILY ledger renderer safe on a protected live sheet.

    Excel does not persist UserInterfaceOnly across sessions.  The renderer therefore
    owns a narrow unprotect/render/reprotect bracket and restores protection on every
    normal or error exit.  No replacement or restore business rule is changed here.
    """

    if "Private Const DAILY_SHEET As String = \"THERAPIST_DAILY\"" not in source:
        raise DailyChangeLedgerLiveProtectionPatchError("Unexpected daily-change-ledger module")

    already = (
        "RenderProtectionExit:" in source
        and "wsDaily.Protect UserInterfaceOnly:=True" in source
        and "wsDaily.Unprotect" in source
    )
    if already:
        return source

    old_setup = '''    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)\n    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)\n    startRow = CLng(wsLedger.Range("R1").Value)\n    If startRow < 2 Then Exit Sub\n'''
    new_setup = '''    Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)\n    Set wsLedger = ThisWorkbook.Worksheets(LEDGER_SHEET)\n    wsDaily.Unprotect\n    startRow = CLng(wsLedger.Range("R1").Value)\n    If startRow < 2 Then GoTo RenderProtectionExit\n'''
    if old_setup not in source:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh setup marker not found")
    source = source.replace(old_setup, new_setup, 1)

    old_empty = "    If lastLedgerRow < 2 Then Exit Sub\n"
    new_empty = "    If lastLedgerRow < 2 Then GoTo RenderProtectionExit\n"
    if old_empty not in source:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger empty-view marker not found")
    source = source.replace(old_empty, new_empty, 1)

    old_tail = '''    wsDaily.Columns("F").ColumnWidth = 18\n    wsDaily.Columns("G").ColumnWidth = 20\n    Exit Sub\n\nRenderError:\n    MsgBox "Δεν ήταν δυνατή η εμφάνιση της λίστας αλλαγών: " & Err.Description, vbExclamation\nEnd Sub\n'''
    new_tail = '''    wsDaily.Columns("F").ColumnWidth = 18\n    wsDaily.Columns("G").ColumnWidth = 20\n\nRenderProtectionExit:\n    On Error Resume Next\n    wsDaily.Protect UserInterfaceOnly:=True\n    On Error GoTo 0\n    Exit Sub\n\nRenderError:\n    Dim renderErrorText As String\n    renderErrorText = Err.Description\n    On Error Resume Next\n    If Not wsDaily Is Nothing Then wsDaily.Protect UserInterfaceOnly:=True\n    On Error GoTo 0\n    MsgBox "Δεν ήταν δυνατή η εμφάνιση της λίστας αλλαγών: " & renderErrorText, vbExclamation\nEnd Sub\n'''
    if old_tail not in source:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh exit marker not found")
    return source.replace(old_tail, new_tail, 1)


def install_daily_change_ledger_live_protection(workbook_path: str | Path) -> None:
    path = Path(workbook_path).resolve()
    if not path.exists():
        raise DailyChangeLedgerLiveProtectionPatchError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise DailyChangeLedgerLiveProtectionPatchError("Workbook must be an .xlsm file")
    if sys.platform != "win32":
        raise DailyChangeLedgerLiveProtectionPatchError("Windows Excel is required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyChangeLedgerLiveProtectionPatchError("pywin32 is required") from exc

    excel = None
    workbook = None
    failure: Exception | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        component = workbook.VBProject.VBComponents(MODULE_NAME)
        code = component.CodeModule
        source = str(code.Lines(1, code.CountOfLines))
        patched = patch_ledger_module_source(source)
        if patched != source:
            code.DeleteLines(1, code.CountOfLines)
            code.AddFromString(patched)
        workbook.Save()
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
        raise DailyChangeLedgerLiveProtectionPatchError(str(failure)) from failure
