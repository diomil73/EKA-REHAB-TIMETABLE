from __future__ import annotations

from pathlib import Path
import sys


MODULE_NAME = "modDailyChangeLedger"


class DailyChangeLedgerLiveProtectionPatchError(RuntimeError):
    pass


def _norm_line(value: str) -> str:
    return " ".join(value.strip().split()).casefold()


def _compact_line(value: str) -> str:
    """Normalize a VBA line for structural matching only.

    VBA permits insignificant whitespace around parentheses and operators, so structural
    matching must not depend on formatting such as Worksheets( DAILY_SHEET ).
    """

    return "".join(value.split()).casefold()


def _find_line(lines: list[str], predicate, *, start: int = 0, end: int | None = None) -> int:
    stop = len(lines) if end is None else min(end, len(lines))
    for index in range(start, stop):
        if predicate(lines[index]):
            return index
    return -1


def patch_ledger_module_source(source: str) -> str:
    """Make the THERAPIST_DAILY ledger renderer safe on a protected live sheet.

    Excel COM can return CRLF source and the live workbook may contain other approved
    edits. This patch therefore works on semantic lines inside RefreshDailyChangeLedgerView
    instead of replacing one exact multi-line block.
    """

    source = source.replace("\r\n", "\n").replace("\r", "\n")
    if "Private Const DAILY_SHEET As String = \"THERAPIST_DAILY\"" not in source:
        raise DailyChangeLedgerLiveProtectionPatchError("Unexpected daily-change-ledger module")

    already = (
        "RenderProtectionExit:" in source
        and "wsDaily.Protect UserInterfaceOnly:=True" in source
        and "wsDaily.Unprotect" in source
    )
    if already:
        return source

    lines = source.split("\n")
    proc_start = _find_line(
        lines,
        lambda line: _compact_line(line).startswith("publicsubrefreshdailychangeledgerview("),
    )
    if proc_start < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh procedure marker not found")

    proc_end = _find_line(lines, lambda line: _compact_line(line) == "endsub", start=proc_start + 1)
    if proc_end < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh procedure end not found")

    error_label = _find_line(lines, lambda line: _compact_line(line) == "rendererror:", start=proc_start + 1, end=proc_end)
    if error_label < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh error label not found")

    on_error = _find_line(
        lines,
        lambda line: _compact_line(line) == "onerrorgotorendererror",
        start=proc_start + 1,
        end=error_label,
    )
    if on_error < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh error handler marker not found")

    if not any(_compact_line(line) == "dimrendererrortextasstring" for line in lines[proc_start:error_label]):
        lines.insert(on_error, "    Dim renderErrorText As String")
        error_label += 1
        proc_end += 1

    set_daily = _find_line(
        lines,
        lambda line: (
            _compact_line(line).startswith("setwsdaily=")
            and "worksheets(daily_sheet)" in _compact_line(line)
        ),
        start=proc_start + 1,
        end=error_label,
    )
    if set_daily < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh DAILY sheet assignment not found")

    if not any(_compact_line(line) == "wsdaily.unprotect" for line in lines[set_daily + 1:error_label]):
        lines.insert(set_daily + 1, "    wsDaily.Unprotect")
        error_label += 1
        proc_end += 1

    start_guard = _find_line(
        lines,
        lambda line: _compact_line(line).startswith("ifstartrow<2then"),
        start=set_daily + 1,
        end=error_label,
    )
    if start_guard >= 0:
        lines[start_guard] = "    If startRow < 2 Then GoTo RenderProtectionExit"

    last_guard = _find_line(
        lines,
        lambda line: _compact_line(line).startswith("iflastledgerrow<2then"),
        start=set_daily + 1,
        end=error_label,
    )
    if last_guard >= 0:
        lines[last_guard] = "    If lastLedgerRow < 2 Then GoTo RenderProtectionExit"

    error_label = _find_line(lines, lambda line: _compact_line(line) == "rendererror:", start=proc_start + 1)
    normal_exit = -1
    for index in range(error_label - 1, proc_start, -1):
        if _compact_line(lines[index]) == "exitsub":
            normal_exit = index
            break
    if normal_exit < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh normal exit not found")

    protection_block = [
        "RenderProtectionExit:",
        "    On Error Resume Next",
        "    wsDaily.Protect UserInterfaceOnly:=True",
        "    On Error GoTo 0",
        "    Exit Sub",
        "",
    ]
    lines[normal_exit:normal_exit + 1] = protection_block

    error_label = _find_line(lines, lambda line: _compact_line(line) == "rendererror:", start=proc_start + 1)
    if error_label < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh error label disappeared")

    error_body = [
        "RenderError:",
        "    renderErrorText = Err.Description",
        "    On Error Resume Next",
        "    If Not wsDaily Is Nothing Then wsDaily.Protect UserInterfaceOnly:=True",
        "    On Error GoTo 0",
    ]
    lines[error_label:error_label + 1] = error_body

    proc_end = _find_line(lines, lambda line: _compact_line(line) == "endsub", start=error_label + len(error_body))
    if proc_end < 0:
        raise DailyChangeLedgerLiveProtectionPatchError("Ledger refresh procedure end disappeared")
    msg_index = _find_line(
        lines,
        lambda line: "msgbox" in _norm_line(line) and "εμφάνιση της λίστας αλλαγών" in line,
        start=error_label + len(error_body),
        end=proc_end,
    )
    if msg_index >= 0:
        lines[msg_index] = '    MsgBox "Δεν ήταν δυνατή η εμφάνιση της λίστας αλλαγών: " & renderErrorText, vbExclamation'

    return "\n".join(lines)


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
        if patched != source.replace("\r\n", "\n").replace("\r", "\n"):
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
