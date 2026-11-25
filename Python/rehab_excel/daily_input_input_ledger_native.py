"""Independent native-Excel read-only INPUT_AND_LEDGER diagnostic.

The existing OOXML saved-state comparator runs first. A *different* read path
then inspects Excel COM-visible DAILY_INPUT/DAILY_CHANGE_LEDGER cells and
reconstructs source facts using the shared canonical business interpreter.
Neither this diagnostic nor its fake-COM tests authenticate verifier provenance,
create a phase receipt, write the workbook, or authorize Gate 2 Apply/Restore.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import sys

from openpyxl import Workbook
from openpyxl.utils.datetime import from_excel

from .authoritative_commit import file_sha256
from .application_contract import DAILY_CHANGE_LEDGER_HEADERS
from .daily_input_applied_state import (
    _source_change_keys, active_input_ledger_keys,
    verify_input_ledger_saved_readback,
)
from .daily_input_day_switch import DailyInputDaySwitchError
from .daily_input_reader import _interpret_daily_input, _parse_date


class InputLedgerNativeReadbackError(RuntimeError):
    """Native Excel independent readback failed closed without PHI in errors."""


@dataclass(frozen=True)
class InputLedgerNativeReadback:
    service_date: date
    workbook_sha256: str
    source_change_count: int
    active_ledger_count: int
    native_excel_version: str


def _cell(ws, row: int, col: int):
    cell = ws.Cells(row, col)
    if bool(cell.HasFormula):
        raise InputLedgerNativeReadbackError("Native Excel source contains formula cell")
    raw = cell.Value2
    if type(raw) is str and raw in {
        "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!",
    }:
        raise InputLedgerNativeReadbackError("Native Excel source contains an error cell")
    return raw


def _source_date(value) -> date:
    if type(value) in (int, float):
        try:
            value = from_excel(value)
        except (ValueError, TypeError, OverflowError) as exc:
            raise InputLedgerNativeReadbackError("Invalid native Excel date") from exc
    try:
        return _parse_date(value)
    except Exception as exc:
        raise InputLedgerNativeReadbackError("Invalid native Excel service date") from exc


def _used_last_row(ws, *, limit: int, kind: str) -> int:
    try:
        area = ws.UsedRange
        last = int(area.Row) + int(area.Rows.Count) - 1
        if last < 1 or last > limit:
            raise InputLedgerNativeReadbackError(f"Native {kind} scope exceeds fixed read limit")
        return last
    except InputLedgerNativeReadbackError:
        raise
    except Exception as exc:
        raise InputLedgerNativeReadbackError(f"Native {kind} used-range invalid") from exc


def _snapshot_native_input(ws, *, limit: int):
    """Copy COM-visible literal values to a private in-memory interpretation grid."""
    scratch = Workbook()
    try:
        target = scratch.active
        target.title = "DAILY_INPUT"
        last = _used_last_row(ws, limit=limit, kind="DAILY_INPUT")
        for row in range(1, last + 1):
            for col in range(1, 7):
                value = _cell(ws, row, col)
                if row == 2 and col == 2:
                    value = _source_date(value)
                if value is not None:
                    target.cell(row, col).value = value
        return scratch
    except Exception:
        scratch.close()
        raise


def _snapshot_native_ledger(ws, *, limit: int):
    """Copy native visible ledger cells to a private in-memory schema worksheet."""
    scratch = Workbook()
    try:
        ledger = scratch.active
        ledger.title = "DAILY_CHANGE_LEDGER"
        last = _used_last_row(ws, limit=limit, kind="ledger")
        if int(ws.Visible) != 2 or not bool(ws.ProtectContents):
            raise InputLedgerNativeReadbackError("Native ledger visibility or protection invalid")
        width = len(DAILY_CHANGE_LEDGER_HEADERS)
        for row in range(1, last + 1):
            for col in range(1, width + 1):
                value = _cell(ws, row, col)
                if value is not None:
                    ledger.cell(row, col).value = value
        return scratch
    except Exception:
        scratch.close()
        raise


def verify_input_ledger_native_readonly(
    workbook_path: str | Path,
    *,
    service_date: date,
    validated_changes: tuple,
    patients,
    sessions,
    therapist_names=(),
    excel_factory=None,
    max_input_row: int = 1500,
    max_ledger_row: int = 10000,
) -> InputLedgerNativeReadback:
    """Require an OOXML readback AND a separate native COM inspection of saved bytes.

    Expected changes originate from the earlier preflight. Never derive them
    from the workbook being verified. The native readback must compare the
    complete source and active ledger identities, not mere counts.
    """
    if type(service_date) is not date:
        raise InputLedgerNativeReadbackError("Explicit service date required")
    if type(validated_changes) is not tuple or not validated_changes:
        raise InputLedgerNativeReadbackError("Non-empty trusted preflight changes required")
    if type(max_input_row) is not int or not 7 <= max_input_row <= 10000:
        raise InputLedgerNativeReadbackError("Invalid input inspection bound")
    if type(max_ledger_row) is not int or not 2 <= max_ledger_row <= 100000:
        raise InputLedgerNativeReadbackError("Invalid ledger inspection bound")
    source = Path(workbook_path).resolve(strict=True)
    if source.suffix.casefold() != ".xlsm" or not source.is_file():
        raise InputLedgerNativeReadbackError("Existing saved XLSM required")
    patients = tuple(patients)
    sessions = tuple(sessions)
    therapist_names = tuple(therapist_names)
    before_hash = file_sha256(source)
    try:
        oo = verify_input_ledger_saved_readback(
            source, service_date=service_date, validated_changes=validated_changes,
            patients=patients, sessions=sessions, therapist_names=therapist_names,
        )
        if oo.workbook_sha256 != before_hash:
            raise InputLedgerNativeReadbackError("OOXML and native source snapshot mismatch")
        expected = _source_change_keys(validated_changes)
    except InputLedgerNativeReadbackError:
        raise
    except Exception as exc:
        raise InputLedgerNativeReadbackError("Independent saved OOXML readback failed") from exc

    if excel_factory is None:
        if sys.platform != "win32":
            raise InputLedgerNativeReadbackError("Native Excel verification requires Windows")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError:
            raise InputLedgerNativeReadbackError("Native Excel COM unavailable") from None
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
        raw_version = getattr(app, "Version", None)
        if raw_version is None or not str(raw_version).strip():
            raise InputLedgerNativeReadbackError("Native Excel version unavailable")
        book = app.Workbooks.Open(
            str(source), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(book.ReadOnly):
            raise InputLedgerNativeReadbackError("Native Excel was not opened read-only")
        native_input = book.Worksheets("DAILY_INPUT")
        native_ledger = book.Worksheets("DAILY_CHANGE_LEDGER")
        input_snapshot = _snapshot_native_input(native_input, limit=max_input_row)
        ledger_snapshot = _snapshot_native_ledger(native_ledger, limit=max_ledger_row)
        try:
            interpreted = _interpret_daily_input(
                input_snapshot.active, patient_list=list(patients),
                session_list=list(sessions), therapist_name_list=therapist_names,
                ignored=set(), allow_empty=False,
            )
            ledger_keys = active_input_ledger_keys(ledger_snapshot.active, service_date)
            if interpreted.target_date != service_date:
                raise InputLedgerNativeReadbackError("Native service date differs from preflight")
            if _source_change_keys(interpreted.source_changes) != expected:
                raise InputLedgerNativeReadbackError("Native DAILY_INPUT differs from preflight")
            if ledger_keys != expected:
                raise InputLedgerNativeReadbackError("Native ledger differs from preflight")
        finally:
            input_snapshot.close()
            ledger_snapshot.close()
        result = InputLedgerNativeReadback(
            service_date=service_date, workbook_sha256=before_hash,
            source_change_count=len(expected), active_ledger_count=len(ledger_keys),
            native_excel_version=str(raw_version).strip(),
        )
    except InputLedgerNativeReadbackError:
        raise
    except Exception as exc:
        raise InputLedgerNativeReadbackError("Independent native Excel readback failed") from exc
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
        raise InputLedgerNativeReadbackError("Native Excel readback cleanup failed")
    if file_sha256(source) != before_hash:
        raise InputLedgerNativeReadbackError("Workbook bytes changed during native readback")
    return result
