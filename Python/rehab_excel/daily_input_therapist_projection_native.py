"""Independent native-Excel read-only THERAPIST_DAILY cell readback.

The complete 126-cell expected grid and both provider-header rows must be
supplied from an independently validated preflight/render plan. This reader
compares Excel-visible values exactly; it does not infer AssignmentIDs from
display labels, and it is not a phase verifier or authorization mechanism.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys

from .authoritative_commit import file_sha256
from .layout_contract import THERAPIST_DAILY_BLOCKS, THERAPIST_DAILY_TIMES, cell_address
from .daily_input_therapist_projection_readback import _clock, _name


class TherapistDailyNativeReadbackError(RuntimeError):
    """Native Excel THERAPIST_DAILY readback failed closed."""


@dataclass(frozen=True)
class TherapistDailyCellExpectation:
    cell: str
    value: str | None


@dataclass(frozen=True)
class TherapistDailyNativeReadback:
    workbook_sha256: str
    checked_provider_headers: int
    checked_grid_cells: int
    occupied_grid_cells: int
    native_excel_version: str


def _expected_grid_cells() -> frozenset[str]:
    return frozenset(
        cell_address(row, col)
        for _header, first, last in THERAPIST_DAILY_BLOCKS
        for row in range(first, last + 1)
        for col in range(2, 11)
    )


def _read_literal(cell):
    try:
        if bool(cell.HasFormula):
            raise TherapistDailyNativeReadbackError("Native therapist grid contains a formula")
        value = cell.Value2
    except TherapistDailyNativeReadbackError:
        raise
    except Exception as exc:
        raise TherapistDailyNativeReadbackError("Native therapist cell cannot be inspected") from exc
    if type(value) is str and value.startswith("#"):
        raise TherapistDailyNativeReadbackError("Native therapist grid contains an Excel error")
    if value is not None and type(value) is not str:
        raise TherapistDailyNativeReadbackError("Native therapist grid contains a non-text patient value")
    return value


def verify_therapist_daily_native_readonly(
    workbook_path: str | Path,
    *,
    expected_provider_headers: tuple[str, ...],
    expected_cells: tuple[TherapistDailyCellExpectation, ...],
    excel_factory=None,
) -> TherapistDailyNativeReadback:
    """Compare the full native provider/time/patient display grid to preflight.

    expected_provider_headers is ordered as B:J for row 1 followed by B:J for
    row 11. expected_cells must contain exactly all 126 patient-grid cells,
    including explicit None/empty expectations. The renderer/preflight must
    supply every value; omitted cells are not silently treated as blank.
    """
    if type(expected_provider_headers) is not tuple or len(expected_provider_headers) != 18:
        raise TherapistDailyNativeReadbackError("Both complete provider header rows are required")
    header_keys = []
    for value in expected_provider_headers:
        if type(value) is not str:
            raise TherapistDailyNativeReadbackError("Provider header expectation is invalid")
        if not value.strip():
            header_keys.append("")
            continue
        try:
            header_keys.append(_name(value))
        except Exception as exc:
            raise TherapistDailyNativeReadbackError("Provider header expectation is invalid") from exc
    for start in (0, 9):
        populated = [key for key in header_keys[start:start + 9] if key]
        if len(set(populated)) != len(populated):
            raise TherapistDailyNativeReadbackError("Duplicate provider in one therapist header block")

    if type(expected_cells) is not tuple or any(
        type(item) is not TherapistDailyCellExpectation
        or type(item.cell) is not str
        or (item.value is not None and type(item.value) is not str)
        for item in expected_cells
    ):
        raise TherapistDailyNativeReadbackError("Immutable complete therapist cell expectations required")
    by_cell = {item.cell.upper(): item.value for item in expected_cells}
    if len(by_cell) != len(expected_cells) or set(by_cell) != _expected_grid_cells():
        raise TherapistDailyNativeReadbackError("Expected therapist grid must cover exactly all 126 cells")

    try:
        source = Path(workbook_path).resolve(strict=True)
        if not source.is_file() or source.suffix.casefold() != ".xlsm":
            raise TherapistDailyNativeReadbackError("Existing saved XLSM required")
        before_hash = file_sha256(source)
    except TherapistDailyNativeReadbackError:
        raise
    except Exception as exc:
        raise TherapistDailyNativeReadbackError("Saved therapist workbook unavailable") from exc

    if excel_factory is None:
        if sys.platform != "win32":
            raise TherapistDailyNativeReadbackError("Native therapist verification requires Windows")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError:
            raise TherapistDailyNativeReadbackError("Native Excel COM unavailable") from None
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    app = None
    book = None
    cleanup_failed = False
    result = None
    try:
        app = excel_factory()
        app.AutomationSecurity = 3
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        version = str(getattr(app, "Version", "") or "").strip()
        if not version:
            raise TherapistDailyNativeReadbackError("Native Excel version unavailable")
        book = app.Workbooks.Open(
            str(source), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(book.ReadOnly):
            raise TherapistDailyNativeReadbackError("Native therapist workbook is not read-only")
        ws = book.Worksheets("THERAPIST_DAILY")

        header_index = 0
        for header_row, first, last in THERAPIST_DAILY_BLOCKS:
            for col in range(2, 11):
                cell = ws.Cells(header_row, col)
                if bool(cell.HasFormula):
                    raise TherapistDailyNativeReadbackError("Native provider header cannot be a formula")
                actual = str(cell.Value or "").strip()
                expected = expected_provider_headers[header_index]
                if expected.strip():
                    if _name(actual) != _name(expected):
                        raise TherapistDailyNativeReadbackError("Native therapist provider header mismatch")
                elif actual:
                    raise TherapistDailyNativeReadbackError("Unexpected provider under a blank header")
                header_index += 1

            for row, expected_time in zip(
                range(first, last + 1), THERAPIST_DAILY_TIMES, strict=True,
            ):
                time_cell = ws.Cells(row, 1)
                if bool(time_cell.HasFormula):
                    raise TherapistDailyNativeReadbackError("Native therapist clock cannot be a formula")
                try:
                    actual_time = _clock(time_cell.Value2)
                except Exception as exc:
                    raise TherapistDailyNativeReadbackError("Native therapist clock is invalid") from exc
                if actual_time != expected_time:
                    raise TherapistDailyNativeReadbackError("Native therapist time grid mismatch")

        occupied = 0
        for address, expected in by_cell.items():
            row_number = int("".join(ch for ch in address if ch.isdigit()))
            column_letters = "".join(ch for ch in address if ch.isalpha())
            column_number = 0
            for char in column_letters:
                column_number = column_number * 26 + ord(char) - ord("A") + 1
            actual = _read_literal(ws.Cells(row_number, column_number))
            if actual != expected:
                raise TherapistDailyNativeReadbackError("Native therapist patient-cell value differs from preflight")
            if actual is not None and actual.strip():
                occupied += 1

        result = TherapistDailyNativeReadback(
            workbook_sha256=before_hash,
            checked_provider_headers=header_index,
            checked_grid_cells=len(by_cell),
            occupied_grid_cells=occupied,
            native_excel_version=version,
        )
    except TherapistDailyNativeReadbackError:
        raise
    except Exception as exc:
        raise TherapistDailyNativeReadbackError("Independent native therapist readback failed") from exc
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
        raise TherapistDailyNativeReadbackError("Native therapist Excel cleanup failed")
    if file_sha256(source) != before_hash:
        raise TherapistDailyNativeReadbackError("Workbook bytes changed during native therapist readback")
    return result
