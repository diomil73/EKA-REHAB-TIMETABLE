"""Native read-only capture of literal DAILY_INPUT source-sheet cells.

The capture compares Excel COM-visible values against independently parsed
OOXML literals for every populated cell in a fixed source-sheet allowlist.
Formula/error cells and ranges beyond fixed bounds fail closed. No writes.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import math
import sys

from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from .authoritative_commit import file_sha256
from .package_probe import WorkbookPackageProbe


class DailyInputSourceNativeCaptureError(RuntimeError):
    """Native source-sheet capture did not match the saved OOXML package."""


# These are source sheets consumed by the DAILY_INPUT projection preflight.
# Required sheets are the minimum workbook contract; the rest are optional
# extension sources whose absence has existing backward-compatible semantics.
SOURCE_SHEET_LIMITS: dict[str, tuple[int, int]] = {
    "PATIENTS": (10000, 64),
    "PATIENT_PLANNER": (10000, 128),
    "SETTINGS": (1000, 32),
    "PROGRAM_WEEKLY_SLOTS": (10000, 16),
    "OUTPATIENT_SCHEDULE": (10000, 32),
    "RECURRING_ASSIGNMENT_IDS": (10000, 8),
    "PATIENT_EXITS": (10000, 16),
    "RECURRING_CHANGE_EVENTS": (10000, 16),
    "STUDENTS": (10000, 16),
    "DAILY_INPUT": (1500, 8),
    "DAILY_CHANGE_LEDGER": (10000, 16),
    "DAILY_REPLACEMENT_FACTS": (10000, 16),
}
REQUIRED_SOURCE_SHEETS = frozenset({
    "PATIENTS", "PATIENT_PLANNER", "SETTINGS",
    "DAILY_INPUT", "DAILY_CHANGE_LEDGER",
})


@dataclass(frozen=True)
class NativeSourceSheetResult:
    sheet_name: str
    populated_cells_checked: int
    used_last_row: int
    used_last_column: int
    formula_cells_checked: int = 0


@dataclass(frozen=True)
class DailyInputSourceNativeCapture:
    workbook_sha256: str
    native_excel_version: str
    sheets: tuple[NativeSourceSheetResult, ...]

    @property
    def checked_sheet_count(self) -> int:
        return len(self.sheets)

    @property
    def checked_populated_cells(self) -> int:
        return sum(item.populated_cells_checked for item in self.sheets)

    @property
    def checked_formula_cells(self) -> int:
        return sum(item.formula_cells_checked for item in self.sheets)


def _planner_formula_source(address: str, formula: str) -> tuple[int, int, str]:
    column_letters, row_number = coordinate_from_string(address)
    if column_letters.upper() not in {"A", "B", "C", "D", "E"}:
        raise DailyInputSourceNativeCaptureError("Unsupported PATIENT_PLANNER formula column")
    expected = f"PATIENTS!{column_letters.upper()}{row_number}"
    if formula.replace("$", "").strip().casefold() != expected.casefold():
        raise DailyInputSourceNativeCaptureError("Unsupported PATIENT_PLANNER formula semantics")
    return row_number, column_index_from_string(column_letters), expected


def _native_cell(ws, address: str):
    column_letters, row_number = coordinate_from_string(address)
    column_number = column_index_from_string(column_letters)
    cell = ws.Cells(row_number, column_number)
    if bool(cell.HasFormula):
        raise DailyInputSourceNativeCaptureError("Native source cell contains a formula")
    value = cell.Value2
    return value


def _same_literal(expected: object, actual: object) -> bool:
    if expected is None or actual is None:
        return expected is None and actual is None
    if type(expected) is bool or type(actual) is bool:
        return type(expected) is bool and type(actual) is bool and expected is actual
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        if not isinstance(actual, (int, float)) or isinstance(actual, bool):
            return False
        return math.isfinite(float(expected)) and math.isfinite(float(actual)) and float(expected) == float(actual)
    return type(expected) is str and type(actual) is str and expected == actual


def _used_range_bounds(ws, *, sheet_name: str) -> tuple[int, int, int, int]:
    try:
        used = ws.UsedRange
        first_row = int(used.Row)
        first_column = int(used.Column)
        last_row = first_row + int(used.Rows.Count) - 1
        last_column = first_column + int(used.Columns.Count) - 1
    except Exception as exc:
        raise DailyInputSourceNativeCaptureError("Native source used range is unavailable") from exc
    row_limit, column_limit = SOURCE_SHEET_LIMITS[sheet_name]
    if (first_row < 1 or first_column < 1 or last_row < first_row
            or last_column < first_column or last_row > row_limit
            or last_column > column_limit):
        raise DailyInputSourceNativeCaptureError("Native source used range exceeds fixed bounds")
    return first_row, first_column, last_row, last_column


def verify_daily_input_source_native_capture_readonly(
    workbook_path: str | Path,
    *,
    excel_factory=None,
) -> DailyInputSourceNativeCapture:
    """Compare all populated source-sheet literals through native Excel COM.

    OOXML formula/error metadata is inspected first. Every literal populated
    cell is then read from a macro-disabled, read-only native Excel session and
    compared without coercing text identifiers into numbers. The source hash
    must remain unchanged. This captures source values, not verifier trust.
    """
    source = Path(workbook_path).resolve(strict=True)
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise DailyInputSourceNativeCaptureError("Existing saved XLSM required")
    before_hash = file_sha256(source)
    probe = WorkbookPackageProbe(source)
    package_sheets = set(probe.sheet_names)
    missing = REQUIRED_SOURCE_SHEETS - package_sheets
    if missing:
        raise DailyInputSourceNativeCaptureError("Required source sheet missing")
    selected = tuple(name for name in SOURCE_SHEET_LIMITS if name in package_sheets)
    package_facts: dict[str, tuple[dict[str, object], int, int]] = {}
    for name in selected:
        formulas = probe.formula_cells(name)
        errors = probe.error_cells(name)
        if errors:
            raise DailyInputSourceNativeCaptureError("Source sheet contains stored Excel error cells")
        if formulas and name != "PATIENT_PLANNER":
            raise DailyInputSourceNativeCaptureError("Source sheet contains unsupported formula cells")
        for address, formula in formulas.items():
            _planner_formula_source(address, formula)
        values = probe.typed_cell_values(name)
        for address in formulas:
            values.pop(address, None)
        for address, value in values.items():
            column_letters, row_number = coordinate_from_string(address)
            column_number = column_index_from_string(column_letters)
            row_limit, column_limit = SOURCE_SHEET_LIMITS[name]
            if row_number > row_limit or column_number > column_limit:
                raise DailyInputSourceNativeCaptureError("OOXML source cell exceeds fixed bounds")
        all_addresses = (*values.keys(), *formulas.keys())
        package_facts[name] = (
            values,
            formulas,
            max((coordinate_from_string(a)[1] for a in all_addresses), default=1),
            max((column_index_from_string(coordinate_from_string(a)[0]) for a in all_addresses), default=1),
        )

    if excel_factory is None:
        if sys.platform != "win32":
            raise DailyInputSourceNativeCaptureError("Native Excel verification requires Windows")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError:
            raise DailyInputSourceNativeCaptureError("Native Excel COM unavailable") from None
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    app = None
    book = None
    cleanup_failed = False
    results: list[NativeSourceSheetResult] = []
    try:
        app = excel_factory()
        app.AutomationSecurity = 3
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        raw_version = getattr(app, "Version", None)
        if raw_version is None or not str(raw_version).strip():
            raise DailyInputSourceNativeCaptureError("Native Excel version unavailable")
        book = app.Workbooks.Open(
            str(source), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(book.ReadOnly):
            raise DailyInputSourceNativeCaptureError("Native Excel was not opened read-only")

        try:
            patients_ws = book.Worksheets("PATIENTS")
        except Exception as exc:
            raise DailyInputSourceNativeCaptureError("Native PATIENTS source worksheet missing") from exc
        for name in selected:
            try:
                ws = book.Worksheets(name)
            except Exception as exc:
                raise DailyInputSourceNativeCaptureError("Native source worksheet missing") from exc
            first_row, first_column, last_row, last_column = _used_range_bounds(ws, sheet_name=name)
            values, formulas, package_last_row, package_last_column = package_facts[name]
            all_addresses = (*values.keys(), *formulas.keys())
            package_first_row = min((coordinate_from_string(a)[1] for a in all_addresses), default=1)
            package_first_column = min((
                column_index_from_string(coordinate_from_string(a)[0]) for a in all_addresses
            ), default=1)
            if (package_first_row < first_row or package_first_column < first_column
                    or package_last_row > last_row or package_last_column > last_column):
                raise DailyInputSourceNativeCaptureError("Native used range does not cover saved source cells")
            for address, expected in values.items():
                actual = _native_cell(ws, address)
                if not _same_literal(expected, actual):
                    raise DailyInputSourceNativeCaptureError("Native source literal differs from saved OOXML")
            for address, formula in formulas.items():
                row_number, column_number, expected_source = _planner_formula_source(address, formula)
                cell = ws.Cells(row_number, column_number)
                if not bool(cell.HasFormula):
                    raise DailyInputSourceNativeCaptureError("Native PATIENT_PLANNER formula is missing")
                actual_formula = str(cell.Formula or "").replace("$", "").lstrip("=").strip().casefold()
                if actual_formula != expected_source.casefold():
                    raise DailyInputSourceNativeCaptureError("Native PATIENT_PLANNER formula linkage mismatch")
                source_cell = patients_ws.Cells(row_number, column_number)
                if bool(source_cell.HasFormula):
                    raise DailyInputSourceNativeCaptureError("Native PATIENTS source cell must be literal")
                source_value = source_cell.Value2
                formula_value = cell.Value2
                if type(source_value) is str and source_value.startswith("#"):
                    raise DailyInputSourceNativeCaptureError("Native PATIENTS source contains an Excel error")
                if not _same_literal(source_value, formula_value):
                    raise DailyInputSourceNativeCaptureError("Native PATIENT_PLANNER formula result differs from PATIENTS")
            results.append(NativeSourceSheetResult(
                sheet_name=name,
                populated_cells_checked=len(values),
                used_last_row=last_row,
                used_last_column=last_column,
                formula_cells_checked=len(formulas),
            ))
    except DailyInputSourceNativeCaptureError:
        raise
    except Exception as exc:
        raise DailyInputSourceNativeCaptureError("Native source-sheet capture failed closed") from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                cleanup_failed = True
        if app is not None:
            try:
                app.Quit()
            except Exception:
                cleanup_failed = True
    if cleanup_failed:
        raise DailyInputSourceNativeCaptureError("Native source capture cleanup failed")
    if file_sha256(source) != before_hash:
        raise DailyInputSourceNativeCaptureError("Workbook bytes changed during native source capture")
    return DailyInputSourceNativeCapture(
        workbook_sha256=before_hash,
        native_excel_version=str(raw_version).strip(),
        sheets=tuple(results),
    )
