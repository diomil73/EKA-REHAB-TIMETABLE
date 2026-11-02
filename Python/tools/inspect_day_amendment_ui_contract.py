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
    DAY_AMENDMENT_MACRO_NAME,
    DAY_AMENDMENT_MODULE_NAME,
)


class DayAmendmentUiInspectionError(RuntimeError):
    pass


def inspect_day_amendment_ui_contract(workbook_path: str | Path) -> dict[str, object]:
    if sys.platform != "win32":
        raise DayAmendmentUiInspectionError("UI inspection requires Windows + Microsoft Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DayAmendmentUiInspectionError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise DayAmendmentUiInspectionError(f"Workbook not found: {path}")

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        if not bool(workbook.HasVBProject):
            raise DayAmendmentUiInspectionError("Workbook has no VBA project")

        try:
            component = workbook.VBProject.VBComponents(DAY_AMENDMENT_MODULE_NAME)
        except Exception as exc:
            raise DayAmendmentUiInspectionError("Day-save VBA module is missing") from exc

        code_module = component.CodeModule
        code = str(code_module.Lines(1, int(code_module.CountOfLines)))

        try:
            button = workbook.Worksheets("DAILY_INPUT").Shapes(DAY_AMENDMENT_BUTTON_NAME)
            button_macro = str(button.OnAction)
        except Exception as exc:
            raise DayAmendmentUiInspectionError("Day-save button is missing from DAILY_INPUT") from exc

        forbidden = {
            "worker": "day_amendment_worker.py" in code,
            "application_quit": "Application.Quit" in code,
            "workbook_close": "ThisWorkbook.Close" in code,
            "close_handoff": "CloseForDaySaveHandoff" in code,
        }
        required = {
            "in_session_prepare": "prepare_day_amendment_in_session.py" in code,
            "in_session_persist": "PersistDaySavePayload" in code,
            "success_message": "Η ημέρα αποθηκεύτηκε επιτυχώς." in code,
        }

        ok = (
            button_macro == DAY_AMENDMENT_MACRO_NAME
            and all(required.values())
            and not any(forbidden.values())
        )
        return {
            "ok": ok,
            "workbook": str(path),
            "button_macro": button_macro,
            "required": required,
            "forbidden": forbidden,
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
    parser = argparse.ArgumentParser(description="Inspect the installed day-save VBA contract without mutating the workbook.")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        result = inspect_day_amendment_ui_contract(args.workbook)
    except Exception as exc:
        print(f"DAY SAVE UI INSPECTION FAILED: {exc}")
        return 2

    print("DAY SAVE UI CONTRACT OK" if result["ok"] else "DAY SAVE UI CONTRACT FAILED")
    print(f"Button macro: {result['button_macro']}")
    for name, value in result["required"].items():
        print(f"Required {name}: {value}")
    for name, value in result["forbidden"].items():
        print(f"Forbidden {name}: {value}")
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
