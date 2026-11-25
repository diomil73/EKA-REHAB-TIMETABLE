"""Read-only native DAILY_INPUT / MASTER / THERAPIST_DAILY diagnostic CLI.

This command never writes the workbook, calls VBA, appends a journal, or
enables Apply/Restore. Shared-slot approvals are not inferred or accepted from
this CLI; if the source has shared provider clocks, preflight fails closed.
"""
from __future__ import annotations

import argparse
from datetime import date
import sys

from rehab_excel.daily_input_projection_preflight import (
    build_daily_input_projection_preflight,
)
from rehab_excel.daily_input_projection_preflight_native import (
    verify_daily_input_projection_preflight_native_readonly,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Read-only native Excel verification of DAILY_INPUT projections",
    )
    parser.add_argument("--source", required=True, help="Authorized disposable XLSM source")
    parser.add_argument("--service-date", required=True, help="Service date YYYY-MM-DD")
    parser.add_argument("--baseline-date", required=True, help="Approved recurring-programme baseline YYYY-MM-DD")
    args = parser.parse_args(argv)

    try:
        service_date = date.fromisoformat(args.service_date)
        baseline_date = date.fromisoformat(args.baseline_date)
    except ValueError:
        print("SAFETY STOP: service-date and baseline-date must be valid YYYY-MM-DD dates")
        return 2

    if sys.platform != "win32":
        print("SAFETY STOP: native Excel COM verification requires Windows")
        return 2

    try:
        preflight = build_daily_input_projection_preflight(
            args.source,
            service_date=service_date,
            baseline_date=baseline_date,
            approved_shared_memberships=(),
        )
        result = verify_daily_input_projection_preflight_native_readonly(
            args.source,
            preflight=preflight,
        )
    except Exception as exc:
        # Deliberately omit exception text: source-reader errors may contain
        # patient/provider identifiers or local workbook details.
        print(f"SAFETY STOP: DAILY_INPUT projection verification failed closed ({type(exc).__name__})")
        return 2

    native = result.native_readback
    print("DAILY_INPUT PROJECTION NATIVE READBACK GREEN")
    print(f"service_date: {result.service_date.isoformat()}")
    print(f"source_workbook_sha256: {result.source_workbook_sha256}")
    print(f"native_source_excel_version: {result.source_capture.native_excel_version}")
    print(f"native_source_sheets_checked: {result.native_source_sheet_count}")
    print(f"native_source_populated_cells_checked: {result.native_source_cells_checked}")
    print(f"patients: {result.patient_count}")
    print(f"stable_base_assignments: {result.base_assignment_count}")
    print(f"dated_sessions: {result.session_count}")
    print(f"daily_input_changes: {result.daily_change_count}")
    print(f"active_replacements: {result.replacement_count}")
    print(f"expected_active_therapist_assignments: {result.therapist_assignment_count}")
    print(f"input_ledger_source_changes: {native.input_ledger_source_changes}")
    print(f"input_ledger_active_rows: {native.input_ledger_active_rows}")
    print(f"master_projected_rows: {native.master_projected_rows}")
    print(f"master_formula_cells_checked: {native.master_formula_cells_checked}")
    print(f"master_evaluated_cells_checked: {native.master_evaluated_cells_checked}")
    print(f"therapist_provider_header_positions_checked: 18")
    print(f"therapist_patient_cells_checked: {native.therapist_grid_cells_checked}")
    print(f"therapist_occupied_cells: {native.therapist_occupied_cells}")
    print(f"master_excel_version: {native.master_excel_version}")
    print(f"therapist_excel_version: {native.therapist_excel_version}")
    print("NON-AUTHORIZING READ-ONLY DIAGNOSTIC ONLY")
    print("No journal append, Apply/Restore, VBA call, workbook save, or Gate 2 authorization.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
