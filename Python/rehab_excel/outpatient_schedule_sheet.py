from __future__ import annotations

from pathlib import Path
import sys


OUTPATIENT_SCHEDULE_SHEET = "OUTPATIENT_SCHEDULE"
OUTPATIENT_SCHEDULE_HEADERS = (
    "PatientID",
    "Ασθενής",
    "Θεραπεία",
    "Ώρα",
    "Ημέρες",
    "Θεραπευτής",
)


class OutpatientScheduleSheetError(RuntimeError):
    pass


def ensure_outpatient_schedule_sheet_in_workbook(workbook_path: str | Path) -> None:
    """Ensure the authoritative outpatient recurring source exists on a safe build copy.

    This helper is build-time only. It never targets the production workbook directly;
    callers are responsible for passing a fresh preview/staging copy.
    """

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise OutpatientScheduleSheetError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise OutpatientScheduleSheetError("Outpatient schedule source requires an .xlsm workbook")
    if sys.platform != "win32":
        raise OutpatientScheduleSheetError("Windows + Microsoft Excel are required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise OutpatientScheduleSheetError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)

        try:
            ws = workbook.Worksheets(OUTPATIENT_SCHEDULE_SHEET)
        except Exception:
            ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
            ws.Name = OUTPATIENT_SCHEDULE_SHEET
            for index, header in enumerate(OUTPATIENT_SCHEDULE_HEADERS, start=1):
                ws.Cells(1, index).Value2 = header
        else:
            current = tuple(
                str(ws.Cells(1, index).Value or "").strip()
                for index in range(1, len(OUTPATIENT_SCHEDULE_HEADERS) + 1)
            )
            if current != OUTPATIENT_SCHEDULE_HEADERS:
                raise OutpatientScheduleSheetError(
                    "OUTPATIENT_SCHEDULE exists with an unexpected header schema"
                )

        verified = tuple(
            str(ws.Cells(1, index).Value or "").strip()
            for index in range(1, len(OUTPATIENT_SCHEDULE_HEADERS) + 1)
        )
        if verified != OUTPATIENT_SCHEDULE_HEADERS:
            raise OutpatientScheduleSheetError(
                "OUTPATIENT_SCHEDULE header verification failed"
            )

        workbook.Save()
    except OutpatientScheduleSheetError:
        raise
    except Exception as exc:
        raise OutpatientScheduleSheetError(
            f"Could not ensure OUTPATIENT_SCHEDULE sheet: {exc}"
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
