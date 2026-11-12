"""One-shot, native Excel installation of opaque IDs on a fresh synthetic CC copy.

No clinical schedule cells are changed. It adds a separate planner identity
registry and column G to weekly/outpatient sources *only if those sources are
unmodified and currently have no AssignmentID metadata*. There is exactly one
native Save on a NEW copy. On any failure freeze the new candidate; never
retry/rewrite it. COM macros are disabled before opening.

This is NOT a released migration tool for operational patient workbooks.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.recurring_change_preview import read_effective_base_entries
from rehab_excel.recurring_id_migration_plan import (
    RecurringIdMigrationPlan, plan_assignment_ids, validate_migration_plan,
)
from rehab_excel.stable_recurring_ids import (
    REGISTRY_SHEET, REGISTRY_HEADERS, COLUMN_NAME, bind_stable_assignment_ids,
)
from tools.prepare_native_probe_copy import create_probe_copy, _sha256
from tools.prove_synthetic_cc_native_write import _session
from tools.verify_h8_synthetic_history import verify_demo_history


FROZEN_CC = "UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm"


class NativeIdInstallError(RuntimeError):
    """Fail closed and keep the failed candidate as evidence."""


def _read_row_text(ws, row: int, col: int) -> str:
    return str(ws.Cells(row, col).Value or "").strip()


def _validate_excel_surface(book, plan: RecurringIdMigrationPlan) -> None:
    """Preflight all affected cells before writing any of them."""
    validate_migration_plan(plan)
    names = {str(s.Name) for s in book.Worksheets}
    if REGISTRY_SHEET in names:
        raise NativeIdInstallError("Identity registry already exists")
    relevant = {r.source_sheet for r in plan.rows if r.source_sheet != "PATIENT_PLANNER"}
    for sheet_name in relevant:
        if sheet_name not in names:
            raise NativeIdInstallError("Recurring source sheet missing")
        ws = book.Worksheets(sheet_name)
        if _read_row_text(ws, 1, 7):
            raise NativeIdInstallError("AssignmentID column already occupied")
        last_g = int(ws.Cells(ws.Rows.Count, 7).End(-4162).Row)  # xlUp
        if last_g > 1:
            raise NativeIdInstallError("Nonempty worksheet column G")
    for row in plan.rows:
        if row.source_row is None:
            continue
        ws = book.Worksheets(row.source_sheet)
        patient_col = 1
        treatment_col = 2 if row.source_sheet == "PROGRAM_WEEKLY_SLOTS" else 3
        if (_read_row_text(ws, row.source_row, patient_col) != row.patient_id
                or _read_row_text(ws, row.source_row, treatment_col) != row.treatment):
            raise NativeIdInstallError("Source row no longer matches immutable ID plan")
        if _read_row_text(ws, row.source_row, 7):
            raise NativeIdInstallError("Source row already has identity metadata")


def _install_identity_cells(book, plan: RecurringIdMigrationPlan) -> None:
    """Only call AFTER an entirely successful preflight."""
    _validate_excel_surface(book, plan)
    new_ws = book.Worksheets.Add()
    new_ws.Name = REGISTRY_SHEET
    for col, heading in enumerate(REGISTRY_HEADERS, start=1):
        new_ws.Cells(1, col).Value = heading
    planner_row = 2
    for row in plan.rows:
        if row.source_sheet == "PATIENT_PLANNER":
            for col, value in enumerate(
                (row.assignment_id, row.patient_id, row.treatment), start=1
            ):
                new_ws.Cells(planner_row, col).Value = value
            planner_row += 1
        else:
            ws = book.Worksheets(row.source_sheet)
            ws.Cells(1, 7).Value = COLUMN_NAME
            ws.Cells(row.source_row, 7).Value = row.assignment_id


def install_native_synthetic_ids(source: Path, output: Path, *, excel_client=None):
    """Exact-source, exclusive-copy, native Save, read-only independent verify."""
    source = Path(source).resolve()
    output = Path(output).resolve()
    if sys.platform != "win32":
        raise NativeIdInstallError("Windows Microsoft Excel is required")
    if source.name != FROZEN_CC or not source.is_file():
        raise NativeIdInstallError("Only frozen synthetic H10 CC source is permitted")
    if (output.exists() or output == source
            or "_DEMO_" not in output.name.upper()
            or output.suffix.lower() != ".xlsm"):
        raise NativeIdInstallError("Fresh, distinct synthetic DEMO .xlsm required")
    if not output.parent.is_dir():
        raise NativeIdInstallError("Output folder missing")
    if excel_client is None:
        import win32com.client as excel_client  # type: ignore[import-not-found]

    source_hash = _sha256(source)
    plan = None
    try:
        verify_demo_history(source)
        entries = read_effective_base_entries(source)
        plan = plan_assignment_ids(entries)
        validate_migration_plan(plan)
        if _sha256(source) != source_hash:
            raise NativeIdInstallError("Source changed during migration preflight")
        create_probe_copy(source, output)
        excel = book = None
        try:
            excel, book = _session(excel_client, output, readonly=False)
            _install_identity_cells(book, plan)
            book.Save()
        finally:
            try:
                if book is not None:
                    book.Close(SaveChanges=False)
            finally:
                if excel is not None:
                    excel.Quit()

        # Independent workbook readback by the existing stable-ID reader;
        # no reliance on the temporary COM session or generated preview.
        current_entries = read_effective_base_entries(output)
        if tuple(current_entries) != tuple(entries):
            raise NativeIdInstallError("Existing treatment programme changed")
        rebound = bind_stable_assignment_ids(output, tuple(current_entries))
        expected = {r.old_source_id: r.assignment_id for r in plan.rows}
        if len(rebound) != len(entries):
            raise NativeIdInstallError("Persisted row count mismatch")
        if any(item.base_entry_id != expected[original.base_entry_id]
               for original, item in zip(entries, rebound)):
            raise NativeIdInstallError("Persisted AssignmentID mismatch")
        verify_demo_history(output)
        with ZipFile(output) as archive:
            if "xl/vbaProject.bin" not in archive.namelist():
                raise NativeIdInstallError("VBA project missing after native Save")
        if _sha256(source) != source_hash:
            raise NativeIdInstallError("Original CC changed")
        return len(plan.rows)
    except Exception as exc:
        # Existing output? Never delete or overwrite. Failed fresh copy is
        # intentionally frozen for forensic inspection.
        raise NativeIdInstallError("Synthetic native ID migration failed; freeze candidate") from exc
    finally:
        if source.is_file() and _sha256(source) != source_hash:
            raise NativeIdInstallError("Original CC changed during migration")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="One-shot IDs on fresh SYNTHETIC H10 CC only")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        count = install_native_synthetic_ids(args.source, args.output)
    except Exception as exc:
        # Do not disclose workbook paths or patient data in shared logs.
        print(f"SAFETY STOP: Native synthetic assignment migration ({type(exc).__name__}).")
        return 2
    print("SYNTHETIC NATIVE ASSIGNMENT IDS GREEN")
    print(f"Verified immutable assignments: {count}")
    print("Original CC unchanged: YES")
    print("Historical synthetic records independently verified: YES")
    print("This is not permission for operational patient migration.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
