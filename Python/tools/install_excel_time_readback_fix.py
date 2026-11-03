from __future__ import annotations

import argparse
from pathlib import Path
import sys


class ExcelTimeReadbackFixError(RuntimeError):
    pass


MODULES = ("modCloseDayAction", "modDayAmendmentAction")
OLD_LINE = (
    '    If IsIsoTimeText(expectedText) And IsDate(actualValue) Then '
    'CellMatchesPayload = (Format$(CDate(actualValue), "hh:nn:ss") = expectedText): Exit Function'
)
NEW_BLOCK = (
    '    If IsIsoTimeText(expectedText) Then\r\n'
    '        If IsDate(actualValue) Then\r\n'
    '            CellMatchesPayload = (Format$(CDate(actualValue), "hh:nn:ss") = expectedText)\r\n'
    '            Exit Function\r\n'
    '        End If\r\n'
    '        If IsNumeric(actualValue) Then\r\n'
    '            CellMatchesPayload = (Format$(CDbl(actualValue), "hh:nn:ss") = expectedText)\r\n'
    '            Exit Function\r\n'
    '        End If\r\n'
    '    End If'
)


def _module_text(code_module) -> str:
    count = int(code_module.CountOfLines)
    if count <= 0:
        return ""
    return str(code_module.Lines(1, count))


def _patch_module(workbook, module_name: str) -> None:
    try:
        component = workbook.VBProject.VBComponents(module_name)
    except Exception as exc:
        raise ExcelTimeReadbackFixError(f"VBA module not found: {module_name}") from exc

    code_module = component.CodeModule
    source = _module_text(code_module)
    count = source.count(OLD_LINE)
    if count != 1:
        raise ExcelTimeReadbackFixError(
            f"Expected exactly one Excel-time comparator target in {module_name}, found {count}"
        )
    source = source.replace(OLD_LINE, NEW_BLOCK, 1)
    code_module.DeleteLines(1, int(code_module.CountOfLines))
    code_module.AddFromString(source)


def _verify(workbook) -> None:
    required = (
        'If IsIsoTimeText(expectedText) Then',
        'If IsNumeric(actualValue) Then',
        'Format$(CDbl(actualValue), "hh:nn:ss") = expectedText',
    )
    for module_name in MODULES:
        component = workbook.VBProject.VBComponents(module_name)
        source = _module_text(component.CodeModule)
        for marker in required:
            if marker not in source:
                raise ExcelTimeReadbackFixError(
                    f"Persisted Excel-time fix marker missing in {module_name}: {marker}"
                )
        if OLD_LINE in source:
            raise ExcelTimeReadbackFixError(
                f"Legacy Date-only time comparator still present in {module_name}"
            )


def install_excel_time_readback_fix(workbook_path: str | Path) -> dict[str, object]:
    if sys.platform != "win32":
        raise ExcelTimeReadbackFixError("Excel-time fix installation requires Windows + Microsoft Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ExcelTimeReadbackFixError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise ExcelTimeReadbackFixError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise ExcelTimeReadbackFixError("Workbook must be an .xlsm file")

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        for module_name in MODULES:
            _patch_module(workbook, module_name)
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
        excel.Quit()
        excel = None

        # Reopen from disk. The installer is green only if the persisted VBA
        # contains the numeric-serial time comparison in both authoritative paths.
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        _verify(workbook)
        return {"ok": True, "workbook": str(path), "modules": MODULES}
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
        description=(
            "Patch close-day and day-save VBA semantic readback so Excel numeric "
            "time serials compare correctly with HH:MM:SS payload values."
        )
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        result = install_excel_time_readback_fix(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("EXCEL TIME READBACK FIX INSTALL OK")
    print(f"Workbook: {result['workbook']}")
    print("Close-day comparator: numeric Excel time serial supported")
    print("Day-save comparator: numeric Excel time serial supported")
    print("Persisted VBA reopen verification: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
