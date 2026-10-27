from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import sys
import tempfile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_input_live_protection_vba import (  # noqa: E402
    DAILY_INPUT_LEDGER_MODULE,
    MODULE_NAME,
    REPLACEMENT_OPERATOR_MODULE,
    install_daily_input_live_protection,
)


class LiveProtectionRepairError(RuntimeError):
    pass


def _component_code(vbproject, name: str) -> str:
    component = vbproject.VBComponents(name)
    module = component.CodeModule
    return str(module.Lines(1, module.CountOfLines))


def _verify(vbproject) -> None:
    module_code = _component_code(vbproject, MODULE_NAME)
    ledger_code = _component_code(vbproject, DAILY_INPUT_LEDGER_MODULE)
    operator_code = _component_code(vbproject, REPLACEMENT_OPERATOR_MODULE)
    if "Public Function RefreshDailyInputProtectionFromLedger() As Boolean" not in module_code:
        raise LiveProtectionRepairError("Live protection function missing after install")
    if "If Not RefreshDailyInputProtectionFromLedger() Then _" not in ledger_code:
        raise LiveProtectionRepairError("Ledger live-protection hook missing after install")
    if '"operator projection protection"' not in operator_code:
        raise LiveProtectionRepairError("Operator live-protection hook missing after install")


def repair_workbook(path: Path) -> None:
    if sys.platform != "win32":
        raise LiveProtectionRepairError("Windows + Microsoft Excel are required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise LiveProtectionRepairError("pywin32 is required") from exc

    rollback = Path(tempfile.gettempdir()) / f"eka_live_protection_{path.stem}.xlsm"
    rollback.unlink(missing_ok=True)
    shutil.copy2(path, rollback)
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        install_daily_input_live_protection(workbook.VBProject)
        workbook.Save()
        _verify(workbook.VBProject)
    except Exception:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
            workbook = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None
        shutil.copy2(rollback, path)
        raise
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
        rollback.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Install live DAILY_INPUT protection hooks into a disposable preview.")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    path = args.workbook.resolve()
    if not path.exists() or path.suffix.casefold() != ".xlsm":
        print(f"SAFETY STOP: invalid .xlsm workbook: {path}")
        return 2
    try:
        repair_workbook(path)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    print("DAILY_INPUT LIVE PROTECTION REPAIR OK")
    print(f"Workbook: {path}")
    print("Session protection refresh: installed")
    print("ACTIVE-row ledger hook: installed")
    print("Manual projection hook: installed")
    print("Rollback protection: enabled")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
