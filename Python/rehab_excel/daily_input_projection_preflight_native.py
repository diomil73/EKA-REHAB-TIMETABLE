"""One-call non-authorizing native DAILY_INPUT projection diagnostic.

Consumes a previously built OOXML source preflight, then runs the native
INPUT_AND_LEDGER, MASTER and full THERAPIST_DAILY readbacks over the same
saved XLSM bytes. It does not create a report, receipt or authorization.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .authoritative_commit import file_sha256
from .daily_input_input_ledger_native import (
    InputLedgerNativeReadback, verify_input_ledger_native_readonly,
)
from .daily_input_projection_native import (
    DailyInputProjectionNativeReadback, verify_daily_input_projection_native_readonly,
)
from .daily_input_projection_preflight import DailyInputProjectionPreflight


class DailyInputProjectionPreflightNativeError(RuntimeError):
    """The complete non-authorizing native projection diagnostic failed closed."""


@dataclass(frozen=True)
class DailyInputProjectionPreflightNativeReadback:
    preflight: DailyInputProjectionPreflight = field(repr=False)
    native_readback: DailyInputProjectionNativeReadback = field(repr=False)

    @property
    def source_workbook_sha256(self) -> str:
        return self.native_readback.workbook_sha256

    @property
    def service_date(self):
        return self.preflight.service_date

    @property
    def patient_count(self) -> int:
        return self.preflight.patient_count

    @property
    def base_assignment_count(self) -> int:
        return self.preflight.base_assignment_count

    @property
    def session_count(self) -> int:
        return self.preflight.session_count

    @property
    def daily_change_count(self) -> int:
        return self.preflight.daily_change_count

    @property
    def replacement_count(self) -> int:
        return self.preflight.replacement_count

    @property
    def therapist_assignment_count(self) -> int:
        return self.preflight.therapist_assignment_count


def verify_daily_input_projection_preflight_native_readonly(
    workbook_path: str | Path,
    *,
    preflight: DailyInputProjectionPreflight,
    input_ledger_excel_factory=None,
    master_excel_factory=None,
    therapist_excel_factory=None,
) -> DailyInputProjectionPreflightNativeReadback:
    """Run the three native readbacks against one independently built preflight.

    Each verifier opens a fresh Excel instance, forces macros disabled, opens
    ReadOnly and closes without saving. All three results must bind to the
    preflight's exact source hash and service date. This is diagnostic only.
    """
    if type(preflight) is not DailyInputProjectionPreflight:
        raise DailyInputProjectionPreflightNativeError("Independent projection preflight object required")
    try:
        source = Path(workbook_path).resolve(strict=True)
        if not source.is_file() or source.suffix.casefold() != ".xlsm":
            raise DailyInputProjectionPreflightNativeError("Existing saved XLSM required")
        before_hash = file_sha256(source)
    except DailyInputProjectionPreflightNativeError:
        raise
    except Exception as exc:
        raise DailyInputProjectionPreflightNativeError("Projection workbook unavailable") from exc

    if (before_hash != preflight.source_workbook_sha256
            or preflight.expected_therapist_grid.source_workbook_sha256 != before_hash
            or preflight.expected_therapist_grid.service_date != preflight.service_date
            or preflight.daily_input.target_date != preflight.service_date
            or not preflight.validated_changes):
        raise DailyInputProjectionPreflightNativeError("Preflight evidence does not bind to this workbook/date")

    try:
        input_ledger = verify_input_ledger_native_readonly(
            source,
            validated_changes=preflight.validated_changes,
            patients=preflight.patients,
            sessions=preflight.sessions,
            therapist_names=preflight.settings.therapist_names,
            excel_factory=input_ledger_excel_factory,
        )
        if type(input_ledger) is not InputLedgerNativeReadback:
            raise DailyInputProjectionPreflightNativeError("Native INPUT_AND_LEDGER verifier returned invalid result")
        if input_ledger.workbook_sha256 != before_hash or input_ledger.service_date != preflight.service_date:
            raise DailyInputProjectionPreflightNativeError("Native INPUT_AND_LEDGER readback binding mismatch")
        if (input_ledger.source_change_count != len(preflight.validated_changes)
                or input_ledger.active_ledger_count != len(preflight.validated_changes)):
            raise DailyInputProjectionPreflightNativeError("Native INPUT_AND_LEDGER change count differs from preflight")

        projections = verify_daily_input_projection_native_readonly(
            source,
            service_date=preflight.service_date,
            input_ledger_readback=input_ledger,
            expected_projection=preflight.expected_projection,
            expected_projection_source_sha256=preflight.source_workbook_sha256,
            expected_therapist_grid=preflight.expected_therapist_grid,
            master_excel_factory=master_excel_factory,
            therapist_excel_factory=therapist_excel_factory,
        )
        if type(projections) is not DailyInputProjectionNativeReadback:
            raise DailyInputProjectionPreflightNativeError("Composite native projection verifier returned invalid result")
        if projections.workbook_sha256 != before_hash or projections.service_date != preflight.service_date:
            raise DailyInputProjectionPreflightNativeError("Composite native projection binding mismatch")
    except DailyInputProjectionPreflightNativeError:
        raise
    except Exception as exc:
        raise DailyInputProjectionPreflightNativeError("Native DAILY_INPUT projection diagnostic failed closed") from exc

    if file_sha256(source) != before_hash:
        raise DailyInputProjectionPreflightNativeError("Workbook bytes changed during full native projection diagnostic")

    return DailyInputProjectionPreflightNativeReadback(
        preflight=preflight,
        native_readback=projections,
    )
