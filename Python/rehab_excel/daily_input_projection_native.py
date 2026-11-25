"""Non-authorizing composite binding for DAILY_INPUT native projection diagnostics.

This composes independently run INPUT_AND_LEDGER, MASTER and full-grid
THERAPIST_DAILY readbacks over the same exact saved XLSM bytes and service date.
It is not a phase report, trusted attestation, receipt or Apply authorization.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .authoritative_commit import file_sha256
from .daily_input_input_ledger_native import InputLedgerNativeReadback
from .daily_input_master_projection_native import (
    MasterProjectionNativeReadback,
    verify_master_projection_native_readonly,
)
from .daily_input_therapist_projection_expected import TherapistDailyExpectedGrid
from .daily_input_therapist_projection_native import (
    TherapistDailyNativeReadback,
    verify_therapist_daily_native_readonly,
)
from .master_projection import MasterProjectionRow


class DailyInputProjectionNativeReadbackError(RuntimeError):
    """Composite native projection evidence does not share one exact source."""


@dataclass(frozen=True)
class DailyInputProjectionNativeReadback:
    workbook_sha256: str
    service_date: date
    master_projected_rows: int
    therapist_grid_cells_checked: int
    therapist_occupied_cells: int
    expected_therapist_assignments: int
    input_ledger_source_changes: int
    input_ledger_active_rows: int
    master_excel_version: str
    therapist_excel_version: str
    master_formula_cells_checked: int = 0
    master_evaluated_cells_checked: int = 0


def verify_daily_input_projection_native_readonly(
    workbook_path: str | Path,
    *,
    service_date: date,
    input_ledger_readback: InputLedgerNativeReadback,
    expected_projection: tuple[MasterProjectionRow, ...],
    expected_projection_source_sha256: str,
    expected_therapist_grid: TherapistDailyExpectedGrid,
    master_excel_factory=None,
    therapist_excel_factory=None,
) -> DailyInputProjectionNativeReadback:
    """Bind three read-only diagnostics to the same workbook bytes/date.

    The INPUT_AND_LEDGER readback must already have succeeded independently.
    The expected MASTER rows and complete THERAPIST_DAILY grid must originate
    from the same trusted preflight as that readback. All components remain
    caller-supplied evidence objects; this function checks consistency only.
    """
    if type(service_date) is not date:
        raise DailyInputProjectionNativeReadbackError("Explicit service date required")
    if type(input_ledger_readback) is not InputLedgerNativeReadback:
        raise DailyInputProjectionNativeReadbackError("Prior native INPUT_AND_LEDGER evidence required")
    if type(expected_projection) is not tuple or any(
        type(item) is not MasterProjectionRow for item in expected_projection
    ):
        raise DailyInputProjectionNativeReadbackError("Immutable independent MASTER expectations required")
    if type(expected_projection_source_sha256) is not str or (
        len(expected_projection_source_sha256) != 64
        or any(ch not in "0123456789abcdef" for ch in expected_projection_source_sha256)
    ):
        raise DailyInputProjectionNativeReadbackError("Independent MASTER preflight source digest required")
    if type(expected_therapist_grid) is not TherapistDailyExpectedGrid:
        raise DailyInputProjectionNativeReadbackError("Complete expected therapist grid required")
    if input_ledger_readback.service_date != service_date:
        raise DailyInputProjectionNativeReadbackError("INPUT_AND_LEDGER service date mismatch")
    if (type(input_ledger_readback.source_change_count) is not int
            or type(input_ledger_readback.active_ledger_count) is not int
            or input_ledger_readback.source_change_count < 0
            or input_ledger_readback.active_ledger_count != input_ledger_readback.source_change_count
            or type(input_ledger_readback.native_excel_version) is not str
            or not input_ledger_readback.native_excel_version.strip()):
        raise DailyInputProjectionNativeReadbackError("INPUT_AND_LEDGER evidence envelope is inconsistent")
    if expected_therapist_grid.service_date != service_date:
        raise DailyInputProjectionNativeReadbackError("Therapist expected-grid service date mismatch")

    try:
        source = Path(workbook_path).resolve(strict=True)
        if not source.is_file() or source.suffix.casefold() != ".xlsm":
            raise DailyInputProjectionNativeReadbackError("Existing saved XLSM required")
        before_hash = file_sha256(source)
    except DailyInputProjectionNativeReadbackError:
        raise
    except Exception as exc:
        raise DailyInputProjectionNativeReadbackError("Saved projection evidence unavailable") from exc

    if (input_ledger_readback.workbook_sha256 != before_hash
            or expected_projection_source_sha256 != before_hash
            or expected_therapist_grid.source_workbook_sha256 != before_hash):
        raise DailyInputProjectionNativeReadbackError("Preflight/readback evidence belongs to different workbook bytes")

    try:
        master = verify_master_projection_native_readonly(
            source,
            expected_projection=expected_projection,
            excel_factory=master_excel_factory,
        )
        if type(master) is not MasterProjectionNativeReadback:
            raise DailyInputProjectionNativeReadbackError("Native MASTER verifier returned an invalid result")
        if master.workbook_sha256 != before_hash:
            raise DailyInputProjectionNativeReadbackError("MASTER readback workbook digest mismatch")
        expected_master_formula_cells = (
            len(expected_projection) * 11
            - sum(bool(item.responsible_doctor) for item in expected_projection)
        )
        expected_master_evaluated_cells = len(expected_projection) * 11
        if (type(master.projected_rows) is not int
                or master.projected_rows < 0
                or master.projected_rows != len(expected_projection)
                or type(master.formula_cells_checked) is not int
                or type(master.evaluated_cells_checked) is not int
                or master.formula_cells_checked != expected_master_formula_cells
                or master.evaluated_cells_checked != expected_master_evaluated_cells
                or type(master.native_excel_version) is not str
                or not master.native_excel_version.strip()):
            raise DailyInputProjectionNativeReadbackError("MASTER readback scope/version mismatch")

        therapist = verify_therapist_daily_native_readonly(
            source,
            expected_provider_headers=expected_therapist_grid.provider_headers,
            expected_cells=expected_therapist_grid.cells,
            excel_factory=therapist_excel_factory,
        )
        if type(therapist) is not TherapistDailyNativeReadback:
            raise DailyInputProjectionNativeReadbackError("Native therapist verifier returned an invalid result")
        if therapist.workbook_sha256 != before_hash:
            raise DailyInputProjectionNativeReadbackError("Therapist readback workbook digest mismatch")
        if (type(therapist.checked_grid_cells) is not int
                or type(therapist.checked_provider_headers) is not int
                or type(therapist.occupied_grid_cells) is not int
                or type(expected_therapist_grid.occupied_cells) is not int
                or therapist.checked_grid_cells != 126
                or therapist.checked_provider_headers != 18
                or not 0 <= therapist.occupied_grid_cells <= 126
                or therapist.occupied_grid_cells != expected_therapist_grid.occupied_cells
                or type(therapist.native_excel_version) is not str
                or not therapist.native_excel_version.strip()):
            raise DailyInputProjectionNativeReadbackError("Native therapist verifier did not match the complete expected grid")
    except DailyInputProjectionNativeReadbackError:
        raise
    except Exception as exc:
        raise DailyInputProjectionNativeReadbackError("Composite native projection readback failed") from exc

    if file_sha256(source) != before_hash:
        raise DailyInputProjectionNativeReadbackError("Workbook bytes changed during composite projection readback")

    return DailyInputProjectionNativeReadback(
        workbook_sha256=before_hash,
        service_date=service_date,
        master_projected_rows=master.projected_rows,
        therapist_grid_cells_checked=therapist.checked_grid_cells,
        therapist_occupied_cells=therapist.occupied_grid_cells,
        expected_therapist_assignments=len(expected_therapist_grid.expected_assignments),
        input_ledger_source_changes=input_ledger_readback.source_change_count,
        input_ledger_active_rows=input_ledger_readback.active_ledger_count,
        master_excel_version=master.native_excel_version,
        therapist_excel_version=therapist.native_excel_version,
        master_formula_cells_checked=master.formula_cells_checked,
        master_evaluated_cells_checked=master.evaluated_cells_checked,
    )
