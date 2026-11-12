"""Fail-closed, read-only plan for assigning immutable recurring IDs.

This plan has NO workbook-write capability. It is intended only for a new,
verified SYNTHETIC demo copy and a subsequent separate native Excel installer.
It never migrates clinical workbooks, production MASTER, CC, or frozen proofs.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable
from uuid import uuid4

from rehab_core.models import BaseScheduleEntry
from .stable_recurring_ids import is_stable_assignment_id


class RecurringIdMigrationError(ValueError):
    """Migration has ambiguous source or conflicting identity metadata."""


@dataclass(frozen=True)
class AssignmentIdRow:
    assignment_id: str
    old_source_id: str
    patient_id: str
    treatment: str
    source_sheet: str
    source_row: int | None


@dataclass(frozen=True)
class RecurringIdMigrationPlan:
    """In-memory source-bound migration plan with no clinical details logged."""
    rows: tuple[AssignmentIdRow, ...]

    @property
    def count(self) -> int:
        return len(self.rows)


def plan_assignment_ids(
    entries: Iterable[BaseScheduleEntry],
    *,
    mint: Callable[[], str] | None = None,
) -> RecurringIdMigrationPlan:
    """Generate an ID per active base entry, never by current worksheet row.

    IDs are opaque UUID-backed tokens. A fresh plan must never be applied to
    an existing registry; native installer verifies source hashes and cells.
    """
    mint = mint or (lambda: "assignment:" + uuid4().hex)
    source = tuple(entries)
    ids = [item.base_entry_id for item in source]
    if len(ids) != len(set(ids)):
        raise RecurringIdMigrationError("Duplicate legacy assignment source ID")
    used_stable: set[str] = set()
    seen_planner: set[tuple[str, str]] = set()
    planned: list[AssignmentIdRow] = []
    for entry in source:
        old_id = entry.base_entry_id
        parts = old_id.split(":")
        if len(parts) == 3 and parts[0] == "planner":
            namespace = "PATIENT_PLANNER"
            source_row = None  # bound by PatientID/Treatment, never row
            key = (entry.patient_id, entry.treatment)
            if key in seen_planner:
                raise RecurringIdMigrationError("Ambiguous planner patient/treatment")
            seen_planner.add(key)
        elif len(parts) == 2 and parts[0] in {"weekly", "outpatient"}:
            namespace = (
                "PROGRAM_WEEKLY_SLOTS" if parts[0] == "weekly"
                else "OUTPATIENT_SCHEDULE"
            )
            try:
                source_row = int(parts[1])
            except ValueError as exc:
                raise RecurringIdMigrationError("Invalid source row") from exc
            if source_row < 2:
                raise RecurringIdMigrationError("Invalid source row")
        else:
            raise RecurringIdMigrationError("Unexpected source identity namespace")
        if not entry.patient_id.strip() or not entry.treatment.strip():
            raise RecurringIdMigrationError("Incomplete recurring assignment")
        token = mint()
        if not is_stable_assignment_id(token) or token in used_stable:
            raise RecurringIdMigrationError("Invalid or duplicate minted AssignmentID")
        used_stable.add(token)
        planned.append(
            AssignmentIdRow(token, old_id, entry.patient_id,
                            entry.treatment, namespace, source_row)
        )
    return RecurringIdMigrationPlan(tuple(planned))


def validate_migration_plan(plan: RecurringIdMigrationPlan) -> None:
    """Reject hand-crafted or mutated plans before COM makes any change."""
    if len({r.assignment_id for r in plan.rows}) != len(plan.rows):
        raise RecurringIdMigrationError("Duplicate planned immutable ID")
    if len({r.old_source_id for r in plan.rows}) != len(plan.rows):
        raise RecurringIdMigrationError("Duplicate planned legacy ID")
    planner_keys: set[tuple[str, str]] = set()
    for row in plan.rows:
        if not is_stable_assignment_id(row.assignment_id):
            raise RecurringIdMigrationError("Malformed planned assignment ID")
        if row.source_sheet == "PATIENT_PLANNER":
            if row.source_row is not None:
                raise RecurringIdMigrationError("Planner may not be row-bound")
            key = (row.patient_id, row.treatment)
            if key in planner_keys:
                raise RecurringIdMigrationError("Duplicate planner treatment key")
            planner_keys.add(key)
        elif row.source_sheet in {"PROGRAM_WEEKLY_SLOTS", "OUTPATIENT_SCHEDULE"}:
            if not isinstance(row.source_row, int) or row.source_row < 2:
                raise RecurringIdMigrationError("Invalid planned worksheet row")
        else:
            raise RecurringIdMigrationError("Unknown migration worksheet")
