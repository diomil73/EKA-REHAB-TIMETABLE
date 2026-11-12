"""Synthetic migration preflight: no workbook access, no live patient data."""
from datetime import time

import pytest

from rehab_core.models import BaseScheduleEntry
from rehab_excel.recurring_id_migration_plan import (
    AssignmentIdRow,
    RecurringIdMigrationError,
    RecurringIdMigrationPlan,
    plan_assignment_ids,
    validate_migration_plan,
)


def entry(source, patient="SYN-P", treatment="ΦΘ"):
    return BaseScheduleEntry(
        source, patient, treatment, time(9, 15), "Δε", "SYN-T",
    )


def ids():
    serial = iter([
        "assignment:SYNTHETIC00000001",
        "assignment:SYNTHETIC00000002",
        "assignment:SYNTHETIC00000003",
    ])
    return lambda: next(serial)


def test_plan_handles_planner_weekly_outpatient_without_mutation():
    entries = (
        entry("planner:18:ΦΘ", treatment="ΦΘ"),
        entry("weekly:4", treatment="Εργο"),
        entry("outpatient:3", patient="SYN-O", treatment="ΦΘ"),
    )
    original = tuple(entries)
    plan = plan_assignment_ids(entries, mint=ids())
    validate_migration_plan(plan)
    assert plan.count == 3
    assert [row.source_sheet for row in plan.rows] == [
        "PATIENT_PLANNER", "PROGRAM_WEEKLY_SLOTS", "OUTPATIENT_SCHEDULE",
    ]
    assert plan.rows[0].source_row is None
    assert [row.source_row for row in plan.rows[1:]] == [4, 3]
    assert entries == original


def test_planner_row_move_is_not_part_of_new_identity():
    first = plan_assignment_ids((entry("planner:18:ΦΘ"),), mint=lambda: "assignment:SYNTHETIC00000001")
    second = plan_assignment_ids((entry("planner:72:ΦΘ"),), mint=lambda: "assignment:SYNTHETIC00000001")
    assert first.rows[0].assignment_id == second.rows[0].assignment_id
    assert first.rows[0].source_row is second.rows[0].source_row is None


@pytest.mark.parametrize("entries", [
    (entry("weekly:2"), entry("weekly:2", patient="OTHER")),
    (entry("planner:2:ΦΘ"), entry("planner:3:ΦΘ")),
    (entry("weekly:1"),),
    (entry("weekly:INVALID"),),
    (entry("unknown:2"),),
])
def test_migration_plan_rejects_ambiguous_or_invalid_source(entries):
    with pytest.raises(RecurringIdMigrationError):
        plan_assignment_ids(entries, mint=ids())


def test_duplicate_minted_id_rejected():
    with pytest.raises(RecurringIdMigrationError, match="duplicate"):
        plan_assignment_ids(
            (entry("weekly:2"), entry("weekly:3")),
            mint=lambda: "assignment:SYNTHETIC00000001",
        )


def test_invalid_minted_id_rejected():
    with pytest.raises(RecurringIdMigrationError, match="Invalid"):
        plan_assignment_ids((entry("weekly:2"),), mint=lambda: "weekly:2")


def test_handcrafted_migration_plan_cannot_point_planner_to_row():
    plan = RecurringIdMigrationPlan((
        AssignmentIdRow("assignment:SYNTHETIC00000001", "planner:2:ΦΘ",
                        "P", "ΦΘ", "PATIENT_PLANNER", 2),
    ))
    with pytest.raises(RecurringIdMigrationError, match="row-bound"):
        validate_migration_plan(plan)


def test_handcrafted_migration_plan_cannot_duplicate_same_target():
    plan = RecurringIdMigrationPlan((
        AssignmentIdRow("assignment:SYNTHETIC00000001", "planner:2:ΦΘ",
                        "P", "ΦΘ", "PATIENT_PLANNER", None),
        AssignmentIdRow("assignment:SYNTHETIC00000002", "planner:3:ΦΘ",
                        "P", "ΦΘ", "PATIENT_PLANNER", None),
    ))
    with pytest.raises(RecurringIdMigrationError, match="Duplicate planner"):
        validate_migration_plan(plan)
