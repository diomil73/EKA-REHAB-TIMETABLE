from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.therapist_history_report_vba import (  # noqa: E402
    THERAPIST_HISTORY_MODULE_NAME,
    THERAPIST_HISTORY_NAV_BUTTON,
    THERAPIST_HISTORY_OPEN_MACRO,
    THERAPIST_HISTORY_SHEET_NAME,
    install_therapist_history_report,
)


class TherapistHistoryUiInstallError(RuntimeError):
    pass


def _audit_persisted(workbook) -> None:
    if not bool(workbook.HasVBProject):
        raise TherapistHistoryUiInstallError("Persisted workbook has no VBA project")
    try:
        _ = workbook.VBProject.VBComponents(THERAPIST_HISTORY_MODULE_NAME)
    except Exception as exc:
        raise TherapistHistoryUiInstallError("Therapist-history VBA module was not persisted") from exc

    try:
        report_ws = workbook.Worksheets(THERAPIST_HISTORY_SHEET_NAME)
        if int(report_ws.Visible) != -1:
            raise TherapistHistoryUiInstallError("Therapist-history report sheet is not visible")
        if not bool(report_ws.ProtectContents):
            raise TherapistHistoryUiInstallError("Therapist-history report contents are not protected")
        if bool(report_ws.ProtectDrawingObjects):
            raise TherapistHistoryUiInstallError("Therapist-history report buttons must remain clickable")
    except TherapistHistoryUiInstallError:
        raise
    except Exception as exc:
        raise TherapistHistoryUiInstallError("Therapist-history report sheet was not confirmed") from exc

    try:
        nav = workbook.Worksheets("THERAPIST_DAILY").Shapes(THERAPIST_HISTORY_NAV_BUTTON)
        if str(nav.OnAction) != THERAPIST_HISTORY_OPEN_MACRO or not bool(nav.Visible):
            raise TherapistHistoryUiInstallError("Therapist-history navigation binding is invalid")
    except TherapistHistoryUiInstallError:
        raise
    except Exception as exc:
        raise TherapistHistoryUiInstallError("Therapist-history navigation button was not confirmed") from exc

    for technical_name in ("FINALIZED_DAY_SNAPSHOTS", "THERAPIST_DAY_HISTORY", "DAY_HISTORY_EVENTS"):
        try:
            technical = workbook.Worksheets(technical_name)
        except Exception as exc:
            raise TherapistHistoryUiInstallError(f"Missing technical history sheet: {technical_name}") from exc
        if int(technical.Visible) != 2:
            raise TherapistHistoryUiInstallError(f"Technical history sheet must remain veryHidden: {technical_name}")


def install_therapist_history_ui(workbook_path: str | Path) -> dict[str, object]:
    if sys.platform != "win32":
        raise TherapistHistoryUiInstallError("Therapist-history UI installation requires Windows + Microsoft Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise TherapistHistoryUiInstallError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise TherapistHistoryUiInstallError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise TherapistHistoryUiInstallError("Workbook must be an .xlsm file")

    result = {
        "ok": True,
        "workbook": str(path),
        "module": THERAPIST_HISTORY_MODULE_NAME,
        "sheet": THERAPIST_HISTORY_SHEET_NAME,
        "button": THERAPIST_HISTORY_NAV_BUTTON,
    }

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        if bool(workbook.ReadOnly):
            raise TherapistHistoryUiInstallError("Workbook opened read-only")
        if not bool(workbook.HasVBProject):
            raise TherapistHistoryUiInstallError("Workbook has no VBA project")

        install_therapist_history_report(workbook.VBProject, workbook)
        _audit_persisted(workbook)
        workbook.Save()
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

    # Verify the saved state in a separate read-only Excel session.
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        if not bool(workbook.ReadOnly):
            raise TherapistHistoryUiInstallError("Persisted verification workbook is not read-only")
        _audit_persisted(workbook)
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

    return result

def main() -> int:
    parser = argparse.ArgumentParser(description="Install therapist-history report UI into an XLSM preview.")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        result = install_therapist_history_ui(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("THERAPIST HISTORY UI INSTALL OK")
    print(f"Workbook: {result['workbook']}")
    print(f"Module: {result['module']}")
    print(f"Sheet: {result['sheet']}")
    print("Entry: THERAPIST_DAILY -> Στατιστικά θεραπευτών")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
