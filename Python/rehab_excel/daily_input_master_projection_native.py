"""Independent native-Excel read-only MASTER projection diagnostic.

This diagnostic layers Excel-evaluated identity readback over the existing
saved-OOXML formula-link audit. Expected MasterProjectionRow facts must come
from an independently validated preflight, never from MASTER_SCHEDULE itself.
It is a diagnostic only: it does not verify therapist projections, assignment
membership, trusted attestation, journal transitions or Apply authorization.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys

from .authoritative_commit import file_sha256
from .daily_input_master_projection_readback import (
    MasterProjectionSavedReadbackError,
    verify_master_projection_saved_readback,
)
from .master_projection import MasterProjectionRow


class MasterProjectionNativeReadbackError(RuntimeError):
    """Native Excel MASTER projection readback failed closed."""


@dataclass(frozen=True)
class MasterProjectionNativeReadback:
    workbook_sha256: str
    projected_rows: int
    header_row: int
    status_column: int
    native_excel_version: str


def _text(value: object) -> str:
    if value is None:
        return ""
    if type(value) is str:
        return value.strip()
    return str(value).strip()


def _formula(cell, expected: str) -> None:
    try:
        if not bool(cell.HasFormula):
            raise MasterProjectionNativeReadbackError("Native MASTER formula missing")
        actual = _text(cell.Formula).replace("$", "").casefold()
    except MasterProjectionNativeReadbackError:
        raise
    except Exception as exc:
        raise MasterProjectionNativeReadbackError("Native MASTER formula cannot be inspected") from exc
    if actual != expected.casefold():
        raise MasterProjectionNativeReadbackError("Native MASTER formula linkage mismatch")


def _value(cell) -> object:
    try:
        if bool(cell.HasFormula) is False:
            raise MasterProjectionNativeReadbackError("Native MASTER expected formula is absent")
        value = cell.Value2
    except MasterProjectionNativeReadbackError:
        raise
    except Exception as exc:
        raise MasterProjectionNativeReadbackError("Native MASTER evaluated value unavailable") from exc
    if type(value) is str and value.startswith("#"):
        raise MasterProjectionNativeReadbackError("Native MASTER contains an Excel error")
    return value


def verify_master_projection_native_readonly(
    workbook_path: str | Path,
    *,
    expected_projection: tuple[MasterProjectionRow, ...],
    excel_factory=None,
) -> MasterProjectionNativeReadback:
    """Compare saved OOXML links and native Excel-evaluated identity fields.

    The native instance is fresh, macro-disabled, read-only, and always closed
    without saving. Only the room, patient display identity and status values
    represented by MasterProjectionRow are compared here. This does not claim
    that all MASTER columns or any therapist projection have been verified.
    """
    if type(expected_projection) is not tuple:
        raise MasterProjectionNativeReadbackError("Immutable independent MASTER expectations required")
    if any(type(item) is not MasterProjectionRow for item in expected_projection):
        raise MasterProjectionNativeReadbackError("Invalid independent MASTER expectation")
    try:
        source = Path(workbook_path).resolve(strict=True)
        if not source.is_file() or source.suffix.casefold() != ".xlsm":
            raise MasterProjectionNativeReadbackError("Existing saved XLSM required")
        before_hash = file_sha256(source)
        verify_master_projection_saved_readback(
            source, expected_projection=expected_projection,
        )
    except MasterProjectionNativeReadbackError:
        raise
    except MasterProjectionSavedReadbackError as exc:
        raise MasterProjectionNativeReadbackError("Saved OOXML MASTER structure failed") from exc
    except Exception as exc:
        raise MasterProjectionNativeReadbackError("Saved MASTER evidence unavailable") from exc

    if excel_factory is None:
        if sys.platform != "win32":
            raise MasterProjectionNativeReadbackError("Native MASTER verification requires Windows")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError:
            raise MasterProjectionNativeReadbackError("Native Excel COM unavailable") from None
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
        version = _text(getattr(app, "Version", None))
        if not version:
            raise MasterProjectionNativeReadbackError("Native Excel version unavailable")
        book = app.Workbooks.Open(
            str(source), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(book.ReadOnly):
            raise MasterProjectionNativeReadbackError("Native Excel workbook is not read-only")
        ws = book.Worksheets("MASTER_SCHEDULE")

        headers = []
        for row in range(1, 11):
            if (_text(ws.Cells(row, 2).Value).casefold() == "θάλαμος".casefold()
                    and _text(ws.Cells(row, 3).Value).casefold() == "ασθενής".casefold()):
                headers.append(row)
        if len(headers) != 1:
            raise MasterProjectionNativeReadbackError("Native MASTER header missing or ambiguous")
        header = headers[0]
        k = _text(ws.Cells(header, 11).Value).casefold()
        l = _text(ws.Cells(header, 12).Value).casefold()
        m = _text(ws.Cells(header, 13).Value).casefold()
        if k == "κατάσταση".casefold():
            status_col = 11
        elif (k == "ψυχολογοι".casefold()
              and l == "απογευματινο προγραμμα".casefold()
              and m == "κατάσταση".casefold()):
            status_col = 13
        else:
            raise MasterProjectionNativeReadbackError("Native MASTER schema unsupported")

        for item in expected_projection:
            row = item.target_row
            _formula(ws.Cells(row, 2), f"=PATIENT_PLANNER!B{item.planner_row}")
            _formula(ws.Cells(row, status_col), f"=PATIENT_PLANNER!E{item.planner_row}")
            expected_identity = item.display_name
            if item.responsible_doctor:
                expected_identity = f"{item.display_name}\n✚ {item.responsible_doctor}"
                cell = ws.Cells(row, 3)
                if bool(cell.HasFormula):
                    raise MasterProjectionNativeReadbackError("Native doctor identity must be a literal")
                actual_identity = _text(cell.Value)
            else:
                _formula(ws.Cells(row, 3), f"=PATIENT_PLANNER!C{item.planner_row}")
                actual_identity = _text(_value(ws.Cells(row, 3)))
            actual_room = _text(_value(ws.Cells(row, 2)))
            actual_status = _text(_value(ws.Cells(row, status_col)))
            if actual_room != item.room.strip():
                raise MasterProjectionNativeReadbackError("Native MASTER room value differs from preflight")
            if actual_identity != expected_identity.strip():
                raise MasterProjectionNativeReadbackError("Native MASTER patient identity differs from preflight")
            if actual_status != _text(item.status):
                raise MasterProjectionNativeReadbackError("Native MASTER status differs from preflight")

        result = MasterProjectionNativeReadback(
            workbook_sha256=before_hash,
            projected_rows=len(expected_projection),
            header_row=header,
            status_column=status_col,
            native_excel_version=version,
        )
    except MasterProjectionNativeReadbackError:
        raise
    except Exception as exc:
        raise MasterProjectionNativeReadbackError("Independent native Excel MASTER readback failed") from exc
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
        raise MasterProjectionNativeReadbackError("Native Excel MASTER cleanup failed")
    if file_sha256(source) != before_hash:
        raise MasterProjectionNativeReadbackError("Workbook bytes changed during native MASTER readback")
    return result
