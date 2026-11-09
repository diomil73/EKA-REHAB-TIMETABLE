from __future__ import annotations

"""Read-only intake audit for a proposed PARALLEL PILOT source workbook.

This tool NEVER converts, installs, repairs, migrates, or approves a workbook.
It prints only counts/status and short safe issue codes, not personal records.
"""

import argparse
from collections import Counter
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile, BadZipFile

from openpyxl import load_workbook

REQUIRED_SHEETS = ("MASTER_SCHEDULE", "DAILY_INPUT", "PATIENTS", "THERAPIST_DAILY")
HISTORY_SHEETS = (
    "FINALIZED_DAY_SNAPSHOTS",
    "THERAPIST_DAY_HISTORY",
    "PATIENT_SESSION_HISTORY",
    "DAY_HISTORY_EVENTS",
    "THERAPIST_ATTENDANCE_HISTORY",
)


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _header_count(ws) -> int:
    if ws.max_row is None or ws.max_row < 2:
        return 0
    return sum(
        1 for (cell,) in ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True)
        if cell is not None and str(cell).strip()
    )


def inspect_source(path: Path) -> dict:
    """Pure inspection: no save operations and no personal values in result."""
    if not path.is_file() or path.suffix.lower() not in (".xlsm", ".xlsx"):
        raise ValueError("Existing .xlsm or .xlsx workbook required")
    before = _digest(path)
    try:
        with ZipFile(path, "r") as archive:
            parts = set(archive.namelist())
            has_vba = any(name.endswith("/vbaProject.bin") for name in parts)
        book = load_workbook(path, read_only=True, data_only=False,
                             keep_vba=path.suffix.lower() == ".xlsm",
                             keep_links=False)
        try:
            present = set(book.sheetnames)
            missing = tuple(name for name in REQUIRED_SHEETS if name not in present)
            history = tuple(name for name in HISTORY_SHEETS if name in present)
            history_counts = {
                name: _header_count(book[name]) for name in history
            }
            has_prior_history = any(history_counts.values())
            registry_rows = 0
            duplicate_ids = 0
            incomplete_registry = 0
            if "PATIENTS" in present:
                registry = book["PATIENTS"]
                seen: set[str] = set()
                for row in registry.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
                    patient_id = row[0]
                    name = row[2]
                    if not any(v is not None and str(v).strip() for v in row):
                        continue
                    registry_rows += 1
                    normalized = str(patient_id).strip().casefold() if patient_id is not None else ""
                    if not normalized or name is None or not str(name).strip():
                        incomplete_registry += 1
                    if normalized:
                        if normalized in seen:
                            duplicate_ids += 1
                        seen.add(normalized)
            issues = []
            if missing:
                issues.append("MISSING_OPERATIONAL_SHEETS")
            if not has_vba:
                issues.append("NO_VBA_PROJECT")
            if path.suffix.lower() != ".xlsm":
                issues.append("NOT_MACRO_ENABLED_XLSM")
            if registry_rows == 0:
                issues.append("EMPTY_OR_MISSING_PATIENT_REGISTRY")
            if incomplete_registry:
                issues.append("INCOMPLETE_PATIENT_REGISTRY_ROWS")
            if duplicate_ids:
                issues.append("DUPLICATE_PATIENT_IDS")
            if has_prior_history:
                issues.append("EXISTING_TEST_OR_UNKNOWN_HISTORY_REQUIRES_ISOLATION")
            return {
                "operational_sheets_found": len(REQUIRED_SHEETS) - len(missing),
                "missing_operational_sheets": missing,
                "history_sheets_found": len(history),
                "history_rows": sum(history_counts.values()),
                "patient_registry_rows": registry_rows,
                "duplicate_patient_ids": duplicate_ids,
                "incomplete_registry_rows": incomplete_registry,
                "vba_present": has_vba,
                "issues": tuple(issues),
            }
        finally:
            book.close()
    except BadZipFile as exc:
        raise ValueError("Invalid Excel ZIP container") from exc
    finally:
        if path.is_file() and _digest(path) != before:
            raise RuntimeError("SAFETY STOP: intake source changed during read-only inspection")


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only parallel pilot workbook intake")
    parser.add_argument("--workbook", type=Path, required=True)
    args = parser.parse_args()
    try:
        results = inspect_source(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    print("PARALLEL PILOT SOURCE READ-ONLY INVENTORY")
    print(f"Operational sheets: {results['operational_sheets_found']}/{len(REQUIRED_SHEETS)}")
    print(f"Patient registry rows: {results['patient_registry_rows']}")
    print(f"Duplicate patient identifiers: {results['duplicate_patient_ids']}")
    print(f"Incomplete registry rows: {results['incomplete_registry_rows']}")
    print(f"Historical sheets: {results['history_sheets_found']}; data rows: {results['history_rows']}")
    print(f"VBA project: {'YES' if results['vba_present'] else 'NO'}")
    for code in results["issues"]:
        print(f"REVIEW REQUIRED: {code}")
    print("Workbook unchanged: True")
    print("NO PILOT APPROVAL, IMPORT, INSTALL, SAVE OR MIGRATION PERFORMED")
    return 3 if results["issues"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
