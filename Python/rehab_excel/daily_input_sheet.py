from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from typing import Iterable, Sequence


SHEET_NAME = "DAILY_INPUT"
LISTS_SHEET_NAME = "_PY_LISTS"


class DailyInputSheetError(RuntimeError):
    pass


def _sha256(path: str | Path) -> str:
    digest = sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _unique_nonblank(values: Iterable[str]) -> tuple[str, ...]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return tuple(result)


def _time_text(value: time) -> str:
    return value.strftime("%H:%M")


def _excel_datetime(value: date) -> datetime:
    """Convert a date-only value to a COM-friendly Excel datetime."""
    return datetime.combine(value, time.min)


@dataclass(frozen=True)
class DailyInputSpec:
    target_date: date
    therapist_names: tuple[str, ...]
    patient_names: tuple[str, ...]
    patient_statuses: tuple[str, ...]
    timeslots: tuple[str, ...]
    yes_no_values: tuple[str, ...] = ("ΝΑΙ", "ΟΧΙ")
    therapist_rows: int = 10
    patient_rows: int = 20

    @property
    def therapist_first_row(self) -> int:
        return 7

    @property
    def therapist_last_row(self) -> int:
        return self.therapist_first_row + self.therapist_rows - 1

    @property
    def patient_header_row(self) -> int:
        return self.therapist_last_row + 4

    @property
    def patient_first_row(self) -> int:
        return self.patient_header_row + 1

    @property
    def patient_last_row(self) -> int:
        return self.patient_first_row + self.patient_rows - 1


def build_daily_input_spec(
    *,
    target_date: date,
    therapist_names: Sequence[str],
    patient_names: Sequence[str],
    patient_statuses: Sequence[str],
    timeslots: Sequence[time],
    therapist_rows: int = 10,
    patient_rows: int = 20,
) -> DailyInputSpec:
    if therapist_rows < 1 or patient_rows < 1:
        raise ValueError("Daily input row counts must be positive")
    return DailyInputSpec(
        target_date=target_date,
        therapist_names=_unique_nonblank(therapist_names),
        patient_names=_unique_nonblank(patient_names),
        patient_statuses=_unique_nonblank(patient_statuses),
        timeslots=tuple(_time_text(value) for value in timeslots),
        therapist_rows=therapist_rows,
        patient_rows=patient_rows,
    )


def _delete_sheet_if_present(workbook, name: str) -> None:
    for sheet in workbook.Worksheets:
        if str(sheet.Name).casefold() == name.casefold():
            sheet.Delete()
            return


def _remove_name_if_present(workbook, name: str) -> None:
    try:
        workbook.Names(name).Delete()
    except Exception:
        pass


def _write_column(ws, column: int, values: Sequence[str]) -> None:
    for row, value in enumerate(values, start=1):
        ws.Cells(row, column).Value = value


def _add_named_range(workbook, name: str, sheet_name: str, column_letter: str, count: int) -> None:
    _remove_name_if_present(workbook, name)
    if count < 1:
        return
    refers_to = f"='{sheet_name}'!${column_letter}$1:${column_letter}${count}"
    workbook.Names.Add(Name=name, RefersTo=refers_to)


def _set_validation(cell_range, formula1: str) -> None:
    # xlValidateList=3, xlValidAlertStop=1, xlBetween=1
    try:
        cell_range.Validation.Delete()
    except Exception:
        pass
    cell_range.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula1)
    cell_range.Validation.IgnoreBlank = True
    cell_range.Validation.InCellDropdown = True
    cell_range.Validation.ShowError = True


def _configure_lists(workbook, spec: DailyInputSpec):
    _delete_sheet_if_present(workbook, LISTS_SHEET_NAME)
    lists = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    lists.Name = LISTS_SHEET_NAME

    _write_column(lists, 1, spec.therapist_names)
    _write_column(lists, 2, spec.patient_names)
    _write_column(lists, 3, spec.patient_statuses)
    _write_column(lists, 4, spec.timeslots)
    _write_column(lists, 5, spec.yes_no_values)

    _add_named_range(workbook, "PY_THERAPISTS", LISTS_SHEET_NAME, "A", len(spec.therapist_names))
    _add_named_range(workbook, "PY_PATIENTS", LISTS_SHEET_NAME, "B", len(spec.patient_names))
    _add_named_range(workbook, "PY_STATUSES", LISTS_SHEET_NAME, "C", len(spec.patient_statuses))
    _add_named_range(workbook, "PY_TIMESLOTS", LISTS_SHEET_NAME, "D", len(spec.timeslots))
    _add_named_range(workbook, "PY_YESNO", LISTS_SHEET_NAME, "E", len(spec.yes_no_values))

    # xlSheetVeryHidden = 2. Users should interact only with DAILY_INPUT.
    lists.Visible = 2
    return lists


def _format_daily_sheet(ws, spec: DailyInputSpec) -> None:
    ws.Cells.Clear()
    ws.Range("A1:F1").Merge()
    ws.Range("A1").Value = "ΚΑΤΑΣΤΑΣΗ ΗΜΕΡΑΣ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").HorizontalAlignment = -4108  # xlCenter

    ws.Range("A2").Value = "Ημερομηνία"
    # Write the date as display text instead of setting NumberFormat.
    # Some localized Excel/COM installations reject Range.NumberFormat even
    # for valid format strings. The DAILY_INPUT date is a UI value, so a
    # deterministic dd/mm/yyyy string is safer and locale-independent.
    ws.Range("B2").Value2 = spec.target_date.strftime("%d/%m/%Y")
    ws.Range("A3:F3").Merge()
    ws.Range("A3").Value = (
        "Συμπλήρωσε μόνο τις γραμμές που χρειάζονται. "
        "Το βασικό πρόγραμμα δεν αλλάζει από αυτή την καρτέλα."
    )
    ws.Range("A3").Font.Italic = True

    ws.Range("A5:F5").Merge()
    ws.Range("A5").Value = "ΑΠΟΥΣΙΕΣ ΘΕΡΑΠΕΥΤΩΝ"
    ws.Range("A5").Font.Bold = True

    therapist_headers = ("Ενεργό", "Θεραπευτής", "Από", "Έως", "Αιτία", "Σχόλιο")
    for col, value in enumerate(therapist_headers, start=1):
        ws.Cells(6, col).Value = value
        ws.Cells(6, col).Font.Bold = True

    patient_title_row = spec.patient_header_row - 1
    ws.Range(f"A{patient_title_row}:F{patient_title_row}").Merge()
    ws.Range(f"A{patient_title_row}").Value = "ΑΠΟΥΣΙΕΣ / ΑΚΥΡΩΣΕΙΣ ΑΣΘΕΝΩΝ"
    ws.Range(f"A{patient_title_row}").Font.Bold = True

    patient_headers = ("Ενεργό", "Ασθενής", "Ώρα", "Όλη ημέρα", "Κατάσταση", "Σχόλιο")
    for col, value in enumerate(patient_headers, start=1):
        ws.Cells(spec.patient_header_row, col).Value = value
        ws.Cells(spec.patient_header_row, col).Font.Bold = True

    # Data validation. All values come from hidden named ranges to avoid locale
    # problems and Excel's 255-character literal-list limit.
    tr1, tr2 = spec.therapist_first_row, spec.therapist_last_row
    pr1, pr2 = spec.patient_first_row, spec.patient_last_row
    _set_validation(ws.Range(f"A{tr1}:A{tr2}"), "=PY_YESNO")
    _set_validation(ws.Range(f"B{tr1}:B{tr2}"), "=PY_THERAPISTS")
    _set_validation(ws.Range(f"C{tr1}:D{tr2}"), "=PY_TIMESLOTS")
    _set_validation(ws.Range(f"A{pr1}:A{pr2}"), "=PY_YESNO")
    _set_validation(ws.Range(f"B{pr1}:B{pr2}"), "=PY_PATIENTS")
    _set_validation(ws.Range(f"C{pr1}:C{pr2}"), "=PY_TIMESLOTS")
    _set_validation(ws.Range(f"D{pr1}:D{pr2}"), "=PY_YESNO")
    _set_validation(ws.Range(f"E{pr1}:E{pr2}"), "=PY_STATUSES")

    # Defaults reduce clicks: a row is inactive unless explicitly turned on.
    ws.Range(f"A{tr1}:A{tr2}").Value = "ΟΧΙ"
    ws.Range(f"A{pr1}:A{pr2}").Value = "ΟΧΙ"
    ws.Range(f"D{pr1}:D{pr2}").Value = "ΟΧΙ"

    # Compact layout, intentionally no giant columns.
    widths = {"A": 10, "B": 24, "C": 10, "D": 10, "E": 24, "F": 28}
    for col, width in widths.items():
        ws.Columns(col).ColumnWidth = width
    ws.Range(f"A1:F{spec.patient_last_row}").VerticalAlignment = -4108
    ws.Range(f"A1:F{spec.patient_last_row}").WrapText = True
    ws.Rows("1:3").RowHeight = 24

    # Light table borders. No hard-coded colors are required for data meaning.
    for rng in (
        ws.Range(f"A6:F{tr2}"),
        ws.Range(f"A{spec.patient_header_row}:F{pr2}"),
    ):
        for edge in (7, 8, 9, 10, 11, 12):
            try:
                border = rng.Borders(edge)
                border.LineStyle = 1
                border.Weight = 2
            except Exception:
                pass

    ws.Activate()
    try:
        ws.Application.ActiveWindow.SplitRow = 5
        ws.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass


def create_daily_input_sheet(workbook_path: str | Path, spec: DailyInputSpec) -> None:
    if sys.platform != "win32":
        raise DailyInputSheetError("DAILY_INPUT preview requires Windows + Microsoft Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyInputSheetError("pywin32 is required: python -m pip install pywin32") from exc

    workbook_path = Path(workbook_path).resolve()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=False)

        _delete_sheet_if_present(workbook, SHEET_NAME)
        _configure_lists(workbook, spec)
        ws = workbook.Worksheets.Add(Before=workbook.Worksheets(1))
        ws.Name = SHEET_NAME
        _format_daily_sheet(ws, spec)
        workbook.Save()
    except Exception as exc:
        raise DailyInputSheetError(f"Could not create DAILY_INPUT sheet: {exc}") from exc
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


def create_daily_input_preview(
    source_path: str | Path,
    output_path: str | Path,
    spec: DailyInputSpec,
    *,
    overwrite: bool = False,
) -> tuple[Path, bool]:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if source == output:
        raise DailyInputSheetError("Preview path must differ from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise DailyInputSheetError("DAILY_INPUT preview requires .xlsm source/output")
    if output.exists() and not overwrite:
        raise DailyInputSheetError(f"Preview already exists: {output}")

    output.parent.mkdir(parents=True, exist_ok=True)
    before = _sha256(source)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)
    try:
        create_daily_input_sheet(output, spec)
    except Exception:
        output.unlink(missing_ok=True)
        raise
    after = _sha256(source)
    if before != after:
        output.unlink(missing_ok=True)
        raise DailyInputSheetError("Source workbook changed during DAILY_INPUT preview creation")
    return output, True
