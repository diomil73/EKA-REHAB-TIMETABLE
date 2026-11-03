from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.patient_history_report_vba import (  # noqa: E402
    PATIENT_REPORT_MODULE,
    PATIENT_REPORT_SHEET,
    PATIENT_REPORT_NAV,
    PATIENT_REPORT_OPEN,
    install_patient_report_components,
)


class PatientHistoryUiInstallError(RuntimeError):
    pass


def _audit(workbook) -> None:
    code = workbook.VBProject.VBComponents(PATIENT_REPORT_MODULE).CodeModule
    source = code.Lines(1, code.CountOfLines)
    if "Public Sub OpenPatientHistoryReport()" not in source:
        raise PatientHistoryUiInstallError("Persisted patient report VBA missing")
    report = workbook.Worksheets(PATIENT_REPORT_SHEET)
    if int(report.Visible) != -1:
        raise PatientHistoryUiInstallError("Patient report is not visible")
    nav = workbook.Worksheets("DAILY_INPUT").Shapes(PATIENT_REPORT_NAV)
    if str(nav.OnAction).split("!")[-1].strip("'") != PATIENT_REPORT_OPEN:
        raise PatientHistoryUiInstallError("Patient report navigation mismatch")
    if not bool(nav.Visible):
        raise PatientHistoryUiInstallError("Patient report entry is not visible")
    for sheet_name in (
        "FINALIZED_DAY_SNAPSHOTS", "THERAPIST_DAY_HISTORY",
        "PATIENT_SESSION_HISTORY", "DAY_HISTORY_EVENTS",
    ):
        if int(workbook.Worksheets(sheet_name).Visible) != 2:
            raise PatientHistoryUiInstallError(f"{sheet_name} must remain veryHidden")


def install_patient_history_ui(path: Path) -> None:
    if sys.platform != "win32":
        raise PatientHistoryUiInstallError("Windows with native Microsoft Excel required")
    if not path.is_file() or path.suffix.casefold() != ".xlsm":
        raise PatientHistoryUiInstallError("Existing disposable XLSM required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientHistoryUiInstallError("pywin32 required") from exc

    excel = win32com.client.DispatchEx("Excel.Application")
    workbook = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        if workbook.ReadOnly:
            raise PatientHistoryUiInstallError("Workbook opened read-only")
        install_patient_report_components(workbook.VBProject, workbook)
        _audit(workbook)
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
        _audit(workbook)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(description="Install H7 patient history UI in disposable XLSM")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    try:
        install_patient_history_ui(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    print("H7 PATIENT HISTORY UI INSTALL OK")
    print("Entry: DAILY_INPUT -> Στατιστικά ασθενών")
    print("Technical history sheets: veryHidden")
    print("Persisted VBA reopen verification: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
