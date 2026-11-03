from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.close_day_action_vba import CLOSE_DAY_MACRO_NAME  # noqa: E402
from rehab_excel.day_amendment_action_vba import (  # noqa: E402
    DAY_AMENDMENT_BUTTON_NAME,
    DAY_AMENDMENT_MACRO_NAME,
)
from tools.build_unified_preview import UnifiedBuildError, build_unified_preview  # noqa: E402
from tools.install_day_amendment_ui import (  # noqa: E402
    DayAmendmentUiInstallError,
    install_day_amendment_ui,
)


HISTORY_SHEETS = (
    "FINALIZED_DAY_SNAPSHOTS",
    "THERAPIST_DAY_HISTORY",
    "PATIENT_SESSION_HISTORY",
    "DAY_HISTORY_EVENTS",
)
CLOSE_DAY_BUTTON_NAME = "ekaToolbar_CloseDay"


class PatientHistorySmokeBuildError(RuntimeError):
    pass


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _assert_clean_history(path: Path) -> None:
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        present = [name for name in HISTORY_SHEETS if name in wb.sheetnames]
    finally:
        wb.close()
    if present:
        raise PatientHistorySmokeBuildError(
            "Fresh H5 smoke unexpectedly contains pre-existing history sheets: "
            + ", ".join(present)
        )


def _install_and_verify_h5_controls(path: Path) -> None:
    if sys.platform != "win32":
        raise PatientHistorySmokeBuildError(
            "H5 smoke control installation requires Windows + Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientHistorySmokeBuildError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        ws = workbook.Worksheets("DAILY_INPUT")

        # The general day-save installer historically placed its button below the
        # final input row. For the native H5 proof both authoritative actions must
        # be visible without scrolling, so own their smoke-only placement here.
        try:
            ws.Shapes(CLOSE_DAY_BUTTON_NAME).Delete()
        except Exception:
            pass
        try:
            ws.Shapes(DAY_AMENDMENT_BUTTON_NAME).Delete()
        except Exception:
            pass

        close_anchor = ws.Range("A4")
        close_button = ws.Shapes.AddShape(
            5,
            float(close_anchor.Left) + 2,
            float(close_anchor.Top) + 1,
            132,
            25,
        )
        close_button.Name = CLOSE_DAY_BUTTON_NAME
        close_button.OnAction = CLOSE_DAY_MACRO_NAME
        close_button.Placement = 3
        close_button.TextFrame2.TextRange.Text = "Κλείσιμο ημέρας"
        close_button.TextFrame2.TextRange.Font.Name = "Calibri"
        close_button.TextFrame2.TextRange.Font.Size = 10
        close_button.TextFrame2.TextRange.Font.Bold = True
        close_button.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        close_button.TextFrame2.VerticalAnchor = 3

        save_anchor = ws.Range("D4")
        save_button = ws.Shapes.AddShape(
            5,
            float(save_anchor.Left) + 2,
            float(save_anchor.Top) + 1,
            140,
            25,
        )
        save_button.Name = DAY_AMENDMENT_BUTTON_NAME
        save_button.OnAction = DAY_AMENDMENT_MACRO_NAME
        save_button.Placement = 3
        save_button.TextFrame2.TextRange.Text = "Αποθήκευση ημέρας"
        save_button.TextFrame2.TextRange.Font.Name = "Calibri"
        save_button.TextFrame2.TextRange.Font.Size = 10
        save_button.TextFrame2.TextRange.Font.Bold = True
        save_button.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        save_button.TextFrame2.VerticalAnchor = 3

        # Make row 4 a stable control strip above the frozen input area.
        ws.Rows(4).RowHeight = 28

        workbook.Save()

        verified_close = ws.Shapes(CLOSE_DAY_BUTTON_NAME)
        verified_save = ws.Shapes(DAY_AMENDMENT_BUTTON_NAME)
        if str(verified_close.OnAction) != CLOSE_DAY_MACRO_NAME:
            raise PatientHistorySmokeBuildError("Close-day button macro binding is invalid")
        if str(verified_save.OnAction) != DAY_AMENDMENT_MACRO_NAME:
            raise PatientHistorySmokeBuildError("Day-save button macro binding is invalid")
        if not bool(verified_close.Visible) or not bool(verified_save.Visible):
            raise PatientHistorySmokeBuildError("H5 action buttons are not visible")
    except PatientHistorySmokeBuildError:
        raise
    except Exception as exc:
        raise PatientHistorySmokeBuildError(
            f"H5 action control installation/verification failed: {exc}"
        ) from exc
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


def build_patient_history_h5_smoke(
    *,
    source: Path,
    output: Path,
    target_date: date,
) -> Path:
    source = source.resolve()
    output = output.resolve()

    if output.exists():
        raise PatientHistorySmokeBuildError(
            f"Output already exists; fresh smoke required: {output}"
        )
    if source == output:
        raise PatientHistorySmokeBuildError("Output must differ from source")

    try:
        build_unified_preview(
            source,
            output,
            target_date=target_date,
            overwrite=False,
        )
        _assert_clean_history(output)
        install_day_amendment_ui(output)
        _install_and_verify_h5_controls(output)
        _assert_clean_history(output)
    except (UnifiedBuildError, DayAmendmentUiInstallError, PatientHistorySmokeBuildError):
        output.unlink(missing_ok=True)
        raise
    except Exception as exc:
        output.unlink(missing_ok=True)
        raise PatientHistorySmokeBuildError(str(exc)) from exc

    return output


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Build one clean native H5 smoke workbook with current close-day and "
            "same-session day-save VBA, visible verified controls, but no pre-existing history sheets."
        )
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--date", type=_parse_date, required=True)
    args = parser.parse_args()

    try:
        output = build_patient_history_h5_smoke(
            source=args.source,
            output=args.output,
            target_date=args.date,
        )
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("H5 PATIENT HISTORY SMOKE BUILD OK")
    print(f"Workbook: {output}")
    print(f"DAILY_INPUT date: {args.date.isoformat()}")
    print("Pre-existing history sheets: 0")
    print("Close-day action: installed")
    print("Day-save action: installed")
    print("Close-day button: visible + verified")
    print("Day-save button: visible + verified")
    print("Ready for native v1 -> v2 patient-history proof")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
