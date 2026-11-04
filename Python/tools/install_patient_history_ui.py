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



REFRESH_NAV = "ekaToolbar_RefreshAll"
REFRESH_MACRO = "RefreshAll"
REFRESH_CAPTION = "ΑΝΑΝΕΩΣΗ"
REFRESH_WIDTH = 92


def _macro_name(shape) -> str:
    try:
        return str(shape.OnAction).split("!")[-1].strip("'")
    except Exception:
        return ""


def _style_refresh_button(shape, *, left: float, top: float) -> None:
    shape.Left = left
    shape.Top = top
    shape.Width = REFRESH_WIDTH
    shape.Height = 26
    shape.OnAction = REFRESH_MACRO
    shape.Placement = 3
    shape.Visible = True
    shape.Fill.ForeColor.RGB = 15527148
    shape.Line.ForeColor.RGB = 10066329
    shape.TextFrame2.TextRange.Text = REFRESH_CAPTION
    shape.TextFrame2.TextRange.Font.Name = "Calibri"
    shape.TextFrame2.TextRange.Font.Size = 9
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
    shape.TextFrame2.TextRange.Font.Fill.Transparency = 0
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.TextFrame2.VerticalAnchor = 3
    shape.ZOrder(0)


def _normalize_statistics_toolbar(source) -> None:
    exit_button = source.Shapes("ekaToolbar_ExitApp")
    if not bool(exit_button.Visible):
        raise PatientHistoryUiInstallError("THERAPIST_DAILY Save & Exit is not visible")

    refresh_shapes = []
    for index in range(1, int(source.Shapes.Count) + 1):
        candidate = source.Shapes(index)
        if _macro_name(candidate) == REFRESH_MACRO:
            refresh_shapes.append(candidate)

    canonical = None
    try:
        canonical = source.Shapes(REFRESH_NAV)
    except Exception:
        pass

    if canonical is not None:
        if _macro_name(canonical) != REFRESH_MACRO:
            raise PatientHistoryUiInstallError("Canonical Refresh button has unexpected macro")
        for candidate in refresh_shapes:
            if str(candidate.Name) != REFRESH_NAV:
                candidate.Delete()
    else:
        if len(refresh_shapes) != 1:
            raise PatientHistoryUiInstallError(
                "Expected exactly one legacy RefreshAll control before migration"
            )
        legacy = refresh_shapes[0]
        left = float(exit_button.Left) + float(exit_button.Width) + 8
        top = float(exit_button.Top)
        canonical = source.Shapes.AddShape(5, left, top, REFRESH_WIDTH, 26)
        canonical.Name = REFRESH_NAV
        try:
            _style_refresh_button(canonical, left=left, top=top)
            if _macro_name(canonical) != REFRESH_MACRO:
                raise PatientHistoryUiInstallError("New RefreshAll macro binding did not persist")
        except Exception:
            try:
                canonical.Delete()
            except Exception:
                pass
            raise
        legacy.Delete()

    refresh_left = float(exit_button.Left) + float(exit_button.Width) + 8
    refresh_top = float(exit_button.Top)
    _style_refresh_button(canonical, left=refresh_left, top=refresh_top)
    if _macro_name(canonical) != REFRESH_MACRO or not bool(canonical.Visible):
        raise PatientHistoryUiInstallError("RefreshAll toolbar button failed verification")

    therapist = source.Shapes("ekaToolbar_TherapistHistory")
    patient = source.Shapes(PATIENT_REPORT_NAV)
    therapist.Left = float(canonical.Left) + float(canonical.Width) + 8
    therapist.Top = refresh_top
    patient.Left = float(therapist.Left) + float(therapist.Width) + 8
    patient.Top = refresh_top


def _audit(workbook) -> None:
    code = workbook.VBProject.VBComponents(PATIENT_REPORT_MODULE).CodeModule
    source = code.Lines(1, code.CountOfLines)
    if "Public Sub OpenPatientHistoryReport()" not in source:
        raise PatientHistoryUiInstallError("Persisted patient report VBA missing")
    try:
        refresh_code = workbook.VBProject.VBComponents("Module3").CodeModule
        refresh_source = str(refresh_code.Lines(1, refresh_code.CountOfLines))
    except Exception as exc:
        raise PatientHistoryUiInstallError("RefreshAll VBA module is unavailable") from exc
    if "Public Sub RefreshAll()" not in refresh_source:
        raise PatientHistoryUiInstallError("RefreshAll VBA procedure is missing")
    report = workbook.Worksheets(PATIENT_REPORT_SHEET)
    if int(report.Visible) != -1:
        raise PatientHistoryUiInstallError("Patient report is not visible")
    exit_button = workbook.Worksheets("THERAPIST_DAILY").Shapes("ekaToolbar_ExitApp")
    refresh = workbook.Worksheets("THERAPIST_DAILY").Shapes(REFRESH_NAV)
    therapist = workbook.Worksheets("THERAPIST_DAILY").Shapes("ekaToolbar_TherapistHistory")
    nav = workbook.Worksheets("THERAPIST_DAILY").Shapes(PATIENT_REPORT_NAV)
    if str(nav.OnAction).split("!")[-1].strip("'") != PATIENT_REPORT_OPEN:
        raise PatientHistoryUiInstallError("Patient report navigation mismatch")
    if not bool(nav.Visible) or not bool(therapist.Visible):
        raise PatientHistoryUiInstallError("Statistics navigation is not visible")
    if _macro_name(refresh) != REFRESH_MACRO or not bool(refresh.Visible):
        raise PatientHistoryUiInstallError("RefreshAll toolbar binding/visibility mismatch")
    refresh_matches = []
    source = workbook.Worksheets("THERAPIST_DAILY")
    for index in range(1, int(source.Shapes.Count) + 1):
        candidate = source.Shapes(index)
        if _macro_name(candidate) == REFRESH_MACRO:
            refresh_matches.append(candidate)
    if len(refresh_matches) != 1 or str(refresh_matches[0].Name) != REFRESH_NAV:
        raise PatientHistoryUiInstallError("Duplicate or legacy RefreshAll control remains")
    if str(refresh.TextFrame2.TextRange.Text).strip() != REFRESH_CAPTION:
        raise PatientHistoryUiInstallError("RefreshAll caption is not readable")
    if abs(float(refresh.Top) - float(exit_button.Top)) > 1:
        raise PatientHistoryUiInstallError("RefreshAll button is outside main toolbar row")
    if float(refresh.Left) < float(exit_button.Left) + float(exit_button.Width) + 4:
        raise PatientHistoryUiInstallError("RefreshAll overlaps Save & Exit")
    if float(therapist.Left) < float(refresh.Left) + float(refresh.Width) + 4:
        raise PatientHistoryUiInstallError("Therapist statistics overlaps RefreshAll")

    if abs(float(nav.Top) - float(therapist.Top)) > 1:
        raise PatientHistoryUiInstallError("Patient and therapist statistics buttons misaligned")
    if float(nav.Left) < float(therapist.Left) + float(therapist.Width) + 4:
        raise PatientHistoryUiInstallError("Statistics navigation buttons overlap")
    if abs(float(therapist.Top) - float(exit_button.Top)) > 1:
        raise PatientHistoryUiInstallError("Statistics buttons are outside main application toolbar")
    if float(therapist.Left) < float(exit_button.Left) + float(exit_button.Width) + 4:
        raise PatientHistoryUiInstallError("Therapist statistics overlaps Save & Exit")

    try:
        workbook.Worksheets("DAILY_INPUT").Shapes(PATIENT_REPORT_NAV)
    except Exception:
        pass
    else:
        raise PatientHistoryUiInstallError("Legacy patient button remains on DAILY_INPUT")
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
        _normalize_statistics_toolbar(workbook.Worksheets("THERAPIST_DAILY"))
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
    print("Entry: THERAPIST_DAILY -> Στατιστικά θεραπευτών + Στατιστικά ασθενών")
    print("RefreshAll: preserved, styled, and placed on main toolbar")
    print("Statistics buttons: aligned after RefreshAll and non-overlapping")
    print("DAILY_INPUT patient button: absent")
    print("Technical history sheets: veryHidden")
    print("Persisted VBA reopen verification: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
