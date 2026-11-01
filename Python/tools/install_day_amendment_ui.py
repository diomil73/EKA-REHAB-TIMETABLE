from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.day_amendment_action_vba import (  # noqa: E402
    DAY_AMENDMENT_BUTTON_NAME,
    DAY_AMENDMENT_MODULE_NAME,
    install_day_amendment_action,
)


class DayAmendmentUiInstallError(RuntimeError):
    pass


def install_day_amendment_ui(workbook_path: str | Path) -> dict[str, object]:
    if sys.platform != "win32":
        raise DayAmendmentUiInstallError("Day-amendment UI installation requires Windows + Microsoft Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DayAmendmentUiInstallError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise DayAmendmentUiInstallError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise DayAmendmentUiInstallError("Workbook must be an .xlsm file")

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        if not bool(workbook.HasVBProject):
            raise DayAmendmentUiInstallError("Workbook has no VBA project")
        vbproject = workbook.VBProject

        # This installer owns only the day-save action. It must not redesign,
        # clean, or otherwise mutate MASTER/navigation while the wider UI
        # contract remains a separate product decision.
        install_day_amendment_action(vbproject, workbook)

        try:
            _ = vbproject.VBComponents(DAY_AMENDMENT_MODULE_NAME)
        except Exception as exc:
            raise DayAmendmentUiInstallError("Day-amendment VBA module was not confirmed") from exc

        try:
            button = workbook.Worksheets("DAILY_INPUT").Shapes(DAY_AMENDMENT_BUTTON_NAME)
            if str(button.OnAction) != "StartClosedDayAmendment":
                raise DayAmendmentUiInstallError("Day-save button macro binding is invalid")
        except DayAmendmentUiInstallError:
            raise
        except Exception as exc:
            raise DayAmendmentUiInstallError("Day-save DAILY_INPUT button was not confirmed") from exc

        workbook.Save()
        return {
            "ok": True,
            "workbook": str(path),
            "module": DAY_AMENDMENT_MODULE_NAME,
            "button": DAY_AMENDMENT_BUTTON_NAME,
            "button_sheet": "DAILY_INPUT",
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
    parser = argparse.ArgumentParser(description="Install only the day-save UI into an XLSM preview.")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        result = install_day_amendment_ui(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DAY AMENDMENT UI INSTALL OK")
    print(f"Workbook: {result['workbook']}")
    print(f"Module: {result['module']}")
    print(f"Button: {result['button']}")
    print("Button sheet: DAILY_INPUT")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
