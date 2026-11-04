from __future__ import annotations

import argparse
from pathlib import Path
import sys


def _shape(ws, name: str):
    try:
        return ws.Shapes(name)
    except Exception:
        return None


def _module_has(wb, module: str, macro: str) -> bool:
    try:
        cm = wb.VBProject.VBComponents(module).CodeModule
        return macro in str(cm.Lines(1, cm.CountOfLines))
    except Exception:
        return False


def inspect(path: Path) -> int:
    """Audit only. No installer, no Save, and explicit nonzero on incomplete UI."""
    if sys.platform != "win32" or not path.is_file() or path.suffix.lower() != ".xlsm":
        print("SAFETY STOP: existing XLSM and Windows Excel required")
        return 2
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 missing")
        return 2

    excel = None
    wb = None
    errors: list[str] = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        print("H7 INTEGRATED READ-ONLY AUDIT")
        print(f"Workbook: {wb.FullName}")
        print("ReadOnly:", bool(wb.ReadOnly))

        master = wb.Worksheets("MASTER_SCHEDULE")
        daily = wb.Worksheets("DAILY_INPUT")
        therapist_daily = wb.Worksheets("THERAPIST_DAILY")
        exit_button = _shape(master, "ekaToolbar_ExitApp")
        save = _shape(master, "ekaToolbar_SaveOperationalDayMaster")
        if exit_button is None or save is None:
            errors.append("MASTER Save Day or Save & Exit missing")
        else:
            print(
                "MASTER:", "SaveDay", str(save.OnAction),
                "left", round(float(save.Left), 1),
                "top", round(float(save.Top), 1),
                "ExitRight", round(float(exit_button.Left) + float(exit_button.Width), 1),
            )
            if str(save.OnAction).split("!")[-1].strip("'") != "SaveOperationalDay":
                errors.append("MASTER Save Day binding incorrect")
            if not bool(save.Visible):
                errors.append("MASTER Save Day not visible")
            if float(save.Left) < float(exit_button.Left) + float(exit_button.Width) + 4:
                errors.append("MASTER Save Day overlaps Save & Exit")

        for name in (
            "ekaToolbar_CloseDay", "ekaToolbar_DayAmendment",
            "ekaToolbar_SaveOperationalDay", "ekaToolbar_PatientHistory",
        ):
            if _shape(daily, name) is not None:
                errors.append(f"DAILY_INPUT forbidden button: {name}")

        therapist_exit = _shape(therapist_daily, "ekaToolbar_ExitApp")
        refresh = _shape(therapist_daily, "ekaToolbar_RefreshAll")
        t = _shape(therapist_daily, "ekaToolbar_TherapistHistory")
        p = _shape(therapist_daily, "ekaToolbar_PatientHistory")
        if refresh is None:
            errors.append("THERAPIST_DAILY RefreshAll toolbar button missing")
        elif therapist_exit is None:
            errors.append("THERAPIST_DAILY Save & Exit missing")
        else:
            print(
                "THERAPIST_DAILY REFRESH:",
                str(refresh.OnAction),
                "left", round(float(refresh.Left), 1),
                "top", round(float(refresh.Top), 1),
                "width", round(float(refresh.Width), 1),
            )
            if str(refresh.OnAction).split("!")[-1].strip("'") != "RefreshAll":
                errors.append("RefreshAll toolbar binding incorrect")
            if not bool(refresh.Visible):
                errors.append("RefreshAll toolbar button not visible")
            if abs(float(refresh.Top) - float(therapist_exit.Top)) > 1:
                errors.append("RefreshAll not on main toolbar row")
            if float(refresh.Left) < float(therapist_exit.Left) + float(therapist_exit.Width) + 4:
                errors.append("RefreshAll overlaps Save & Exit")

        if t is None or p is None:
            errors.append("THERAPIST_DAILY requires both statistics buttons")
        else:
            print(
                "THERAPIST_DAILY:",
                "Therapist", str(t.OnAction),
                "Patient", str(p.OnAction),
                "Top", round(float(t.Top), 1), round(float(p.Top), 1),
                "Left", round(float(t.Left), 1), round(float(p.Left), 1),
            )
            if str(t.OnAction).split("!")[-1].strip("'") != "OpenTherapistHistoryReport":
                errors.append("Therapist report button binding incorrect")
            if str(p.OnAction).split("!")[-1].strip("'") != "OpenPatientHistoryReport":
                errors.append("Patient report button binding incorrect")
            if not bool(t.Visible) or not bool(p.Visible):
                errors.append("Statistics button is not visible")
            if abs(float(t.Top) - float(p.Top)) > 1:
                errors.append("Statistics buttons not aligned")
            if therapist_exit is None:
                errors.append("THERAPIST_DAILY Save & Exit missing")
            elif (abs(float(t.Top) - float(therapist_exit.Top)) > 1 or
                  float(t.Left) < float(therapist_exit.Left) + float(therapist_exit.Width) + 4):
                errors.append("Statistics buttons not on main navigation row after Save & Exit")
            if refresh is not None and float(t.Left) < float(refresh.Left) + float(refresh.Width) + 4:
                errors.append("Therapist statistics overlaps RefreshAll")
            if float(p.Left) < float(t.Left) + float(t.Width) + 4:
                errors.append("Statistics buttons overlap")

        for module, macro in (
            ("modUnifiedDaySave", "Public Sub SaveOperationalDay()"),
            ("modTherapistHistoryReport", "Public Sub OpenTherapistHistoryReport()"),
            ("modPatientHistoryReport", "Public Sub OpenPatientHistoryReport()"),
        ):
            present = _module_has(wb, module, macro)
            print(f"VBA {module}:", "PRESENT" if present else "MISSING")
            if not present:
                errors.append(f"VBA {module} missing")

        for name in ("THERAPIST_HISTORY_REPORT", "PATIENT_HISTORY_REPORT"):
            try:
                visible = int(wb.Worksheets(name).Visible)
            except Exception:
                visible = None
            print(f"REPORT {name}: visibility={visible}")
            if visible != -1:
                errors.append(f"{name} not visible")

        for name in (
            "FINALIZED_DAY_SNAPSHOTS", "THERAPIST_DAY_HISTORY",
            "PATIENT_SESSION_HISTORY", "DAY_HISTORY_EVENTS",
        ):
            try:
                visible = int(wb.Worksheets(name).Visible)
            except Exception:
                visible = None
            if visible != 2:
                errors.append(f"{name} not veryHidden")
        print("HISTORY SHEETS:", "all veryHidden" if not any("veryHidden" in e for e in errors) else "FAILED")
    except Exception as exc:
        errors.append(f"Inspection failed: {type(exc).__name__}: {exc}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    if errors:
        for error in errors:
            print("AUDIT FAIL:", error)
        print("H7 INTEGRATED AUDIT INCOMPLETE; workbook not changed")
        return 2
    print("H7 INTEGRATED AUDIT OK; workbook not changed")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only H6/H4/H7 persisted workbook audit")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return inspect(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
