from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import sys
from typing import Iterable

from rehab_core.daily_statistics import DailyTreatmentStatistics


SHEET_NAME = "DAILY_STATISTICS"


class DailyStatisticsWriteError(RuntimeError):
    pass


@dataclass(frozen=True)
class DailyStatisticRow:
    key: str
    label: str
    value: int | str


def build_daily_statistics_rows(
    target_date: date,
    statistics: DailyTreatmentStatistics,
) -> tuple[DailyStatisticRow, ...]:
    return (
        DailyStatisticRow("date", "Ημερομηνία", target_date.strftime("%d/%m/%Y")),
        DailyStatisticRow("scheduled", "Προγραμματισμένες συνεδρίες", statistics.scheduled_sessions),
        DailyStatisticRow("performed", "Πραγματοποιήθηκαν", statistics.performed_sessions),
        DailyStatisticRow("replaced", "Με ανακατανομή", statistics.replaced_sessions),
        DailyStatisticRow("patient_absent", "Απουσία ασθενούς", statistics.patient_absent_sessions),
        DailyStatisticRow("department_postponed", "Αναβολή τμήματος", statistics.department_postponed_sessions),
        DailyStatisticRow("patient_no_show", "Δεν προσήλθε", statistics.patient_no_show_sessions),
        DailyStatisticRow("treatment_suspended", "Αναστολή θεραπείας", statistics.treatment_suspended_sessions),
        DailyStatisticRow("unresolved", "Ορφανοί χωρίς τελική απόφαση", statistics.unresolved_sessions),
        DailyStatisticRow("finalized", "Ημέρα οριστικοποιημένη", "ΝΑΙ" if statistics.is_finalized else "ΟΧΙ"),
    )


def write_daily_statistics(
    workbook_path: str | Path,
    target_date: date,
    statistics: DailyTreatmentStatistics,
) -> None:
    if sys.platform != "win32":
        raise DailyStatisticsWriteError(
            "Daily statistics workbook output requires Windows + Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyStatisticsWriteError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    rows = build_daily_statistics_rows(target_date, statistics)
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)

        ws = None
        for sheet in workbook.Worksheets:
            if str(sheet.Name).casefold() == SHEET_NAME.casefold():
                ws = sheet
                break
        if ws is None:
            ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
            ws.Name = SHEET_NAME
        else:
            ws.Cells.Clear()

        ws.Range("A1:C1").Merge()
        ws.Range("A1").Value = "ΗΜΕΡΗΣΙΑ ΣΤΑΤΙΣΤΙΚΑ ΘΕΡΑΠΕΙΩΝ"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
        ws.Range("A3").Value = "Δείκτης"
        ws.Range("B3").Value = "Περιγραφή"
        ws.Range("C3").Value = "Τιμή"
        ws.Range("A3:C3").Font.Bold = True

        for offset, row in enumerate(rows, start=4):
            ws.Cells(offset, 1).Value = row.key
            ws.Cells(offset, 2).Value = row.label
            ws.Cells(offset, 3).Value = row.value

        ws.Columns("A").ColumnWidth = 24
        ws.Columns("B").ColumnWidth = 34
        ws.Columns("C").ColumnWidth = 18
        try:
            ws.Range(f"A3:C{3 + len(rows)}").Borders.LineStyle = 1
        except Exception:
            pass
        workbook.Save()
    except DailyStatisticsWriteError:
        raise
    except Exception as exc:
        raise DailyStatisticsWriteError(f"Could not write daily statistics: {exc}") from exc
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
