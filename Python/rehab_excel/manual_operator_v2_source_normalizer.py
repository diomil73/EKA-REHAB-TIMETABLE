from __future__ import annotations

from pathlib import Path
import sys


MODULE_NAMES = ("modReplacementOperator", "modDailyChangeLedger")


class ManualOperatorV2SourceNormalizerError(RuntimeError):
    pass


def normalize_vba_source(source: str) -> str:
    return str(source).replace("\r\n", "\n").replace("\r", "\n")


def normalize_manual_operator_v2_sources(workbook_path: str | Path) -> None:
    path = Path(workbook_path).resolve()
    if not path.exists():
        raise ManualOperatorV2SourceNormalizerError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise ManualOperatorV2SourceNormalizerError("Workbook must be an .xlsm file")
    if sys.platform != "win32":
        raise ManualOperatorV2SourceNormalizerError("Windows Excel is required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ManualOperatorV2SourceNormalizerError("pywin32 is required") from exc

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

        for module_name in MODULE_NAMES:
            code = workbook.VBProject.VBComponents(module_name).CodeModule
            source = str(code.Lines(1, code.CountOfLines))
            normalized = normalize_vba_source(source)
            # Rewrite through AddFromString so the downstream patcher sees one
            # stable representation regardless of how Excel COM returned Lines().
            code.DeleteLines(1, code.CountOfLines)
            code.AddFromString(normalized)

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
        raise ManualOperatorV2SourceNormalizerError(str(failure)) from failure
