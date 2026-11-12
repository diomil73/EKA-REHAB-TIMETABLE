"""Opt-in immutable IDs for recurring assignments, read-only.

Legacy row-derived IDs remain available for viewing old programmes, but MUST
NOT be accepted for persistent effective-dated change events.

For PATIENT_PLANNER, one treatment per PatientID is the existing contract:
an optional separate identity registry binds (PatientID, Treatment) to its
immutable ID without depending on its row number. Granular weekly rows and
OUTPATIENT_SCHEDULE rows need their own AssignmentID in column seven. Never
infer an ID from patient name, therapist, time, weekday or mutable row index.
"""
from __future__ import annotations

from dataclasses import replace
from pathlib import Path
import re

from openpyxl import load_workbook

from rehab_core.models import BaseScheduleEntry


REGISTRY_SHEET = "RECURRING_ASSIGNMENT_IDS"
REGISTRY_HEADERS = ("AssignmentID", "PatientID", "Treatment")
COLUMN_NAME = "AssignmentID"
VALID_ID = re.compile(r"^assignment:[A-Za-z0-9][A-Za-z0-9_-]{7,79}$")


class StableRecurringIdentityError(ValueError):
    """Cannot prove an immutable assignment identity unambiguously."""


def is_stable_assignment_id(value: str) -> bool:
    return bool(VALID_ID.fullmatch(value))


def _require_id(raw: object) -> str:
    value = str(raw or "").strip()
    if not is_stable_assignment_id(value):
        raise StableRecurringIdentityError("Missing or malformed immutable AssignmentID")
    return value


def bind_stable_assignment_ids(
    workbook_path: str | Path,
    entries: tuple[BaseScheduleEntry, ...],
    *,
    require_complete: bool = True,
) -> tuple[BaseScheduleEntry, ...]:
    """Bind IDs by source ownership, never by current row number as identity.

    Old row indices are used ONLY to read the embedded ID from the source row;
    the ID itself persists if the entire row moves. Planner IDs are bound by
    the unique stable (PatientID,Treatment) key, not a worksheet row.
    """
    path = Path(workbook_path)
    wb = load_workbook(path, read_only=True, data_only=True,
                       keep_vba=path.suffix.lower() == ".xlsm")
    try:
        planner_ids: dict[tuple[str, str], str] = {}
        declared_planner_ids: set[str] = set()
        if REGISTRY_SHEET in wb.sheetnames:
            ws = wb[REGISTRY_SHEET]
            headers = tuple(ws.cell(1, i).value for i in range(1, 4))
            if headers != REGISTRY_HEADERS:
                raise StableRecurringIdentityError("Invalid recurring identity registry headers")
            for values in ws.iter_rows(min_row=2, max_col=3, values_only=True):
                if all(v is None or str(v).strip() == "" for v in values):
                    continue
                identifier = _require_id(values[0])
                patient = str(values[1] or "").strip()
                treatment = str(values[2] or "").strip()
                if not patient or not treatment:
                    raise StableRecurringIdentityError("Incomplete planner identity registry row")
                key = (patient, treatment)
                if key in planner_ids:
                    raise StableRecurringIdentityError("Duplicate planner identity key")
                if identifier in declared_planner_ids:
                    raise StableRecurringIdentityError("Duplicate immutable planner AssignmentID")
                declared_planner_ids.add(identifier)
                planner_ids[key] = identifier

        used_ids: set[str] = set()
        used_planner_keys: set[tuple[str, str]] = set()
        resolved = []
        for entry in entries:
            legacy_id = entry.base_entry_id
            if legacy_id.startswith("planner:"):
                key = (entry.patient_id, entry.treatment)
                if key in used_planner_keys:
                    raise StableRecurringIdentityError(
                        "Duplicate planner assignment key cannot be bound"
                    )
                used_planner_keys.add(key)
                identifier = planner_ids.get(key)
            elif legacy_id.startswith(("weekly:", "outpatient:")):
                namespace = "PROGRAM_WEEKLY_SLOTS" if legacy_id.startswith("weekly:") else "OUTPATIENT_SCHEDULE"
                if namespace not in wb.sheetnames:
                    raise StableRecurringIdentityError("Recurring source sheet missing")
                ws = wb[namespace]
                # Header position may be extended, but existing columns are
                # never rewritten or reordered by this reader.
                header = tuple(c.value for c in ws[1])
                if COLUMN_NAME not in header:
                    identifier = None
                else:
                    col = header.index(COLUMN_NAME) + 1
                    try:
                        source_row = int(legacy_id.rsplit(":", 1)[1])
                    except ValueError as exc:
                        raise StableRecurringIdentityError("Unknown row-derived source key") from exc
                    row_patient = str(ws.cell(source_row, 1).value or "").strip()
                    treatment_col = 2 if namespace == "PROGRAM_WEEKLY_SLOTS" else 3
                    row_treatment = str(ws.cell(source_row, treatment_col).value or "").strip()
                    if row_patient != entry.patient_id or row_treatment != entry.treatment:
                        raise StableRecurringIdentityError(
                            "Source row changed; identity cannot be bound"
                        )
                    identifier = ws.cell(source_row, col).value
            else:
                raise StableRecurringIdentityError("Unknown recurring source namespace")
            if identifier is None:
                if require_complete:
                    raise StableRecurringIdentityError("Recurring assignment lacks immutable ID")
                resolved.append(entry)
                continue
            identity = _require_id(identifier)
            if identity in used_ids:
                raise StableRecurringIdentityError("Duplicate immutable assignment ID")
            used_ids.add(identity)
            resolved.append(replace(entry, base_entry_id=identity))
        if require_complete and len(resolved) != len(used_ids):
            raise StableRecurringIdentityError("Partial assignment identity mapping")
        return tuple(resolved)
    finally:
        wb.close()
