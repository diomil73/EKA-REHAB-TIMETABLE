from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class SheetAuthority(str, Enum):
    """How Python is allowed to interpret a workbook sheet."""

    AUTHORITATIVE = "authoritative"
    DERIVED_VIEW = "derived_view"
    LEGACY_OPERATIONAL = "legacy_operational"
    AUDIT_ONLY = "audit_only"
    UI_ONLY = "ui_only"


@dataclass(frozen=True)
class SheetSourceRule:
    sheet: str
    authority: SheetAuthority
    purpose: str
    python_import: bool = False
    future_write_target: bool = False


# Only rules already confirmed by the project are marked authoritative.
# The other workbook sheets are deliberately classified conservatively until
# their redesigned role is agreed sheet-by-sheet.
SOURCE_RULES: dict[str, SheetSourceRule] = {
    "PATIENTS": SheetSourceRule(
        sheet="PATIENTS",
        authority=SheetAuthority.AUTHORITATIVE,
        purpose="Patient registry and current patient state",
        python_import=True,
    ),
    "PATIENT_PLANNER": SheetSourceRule(
        sheet="PATIENT_PLANNER",
        authority=SheetAuthority.AUTHORITATIVE,
        purpose="Recurring base programme",
        python_import=True,
    ),
    "SETTINGS": SheetSourceRule(
        sheet="SETTINGS",
        authority=SheetAuthority.AUTHORITATIVE,
        purpose="Workbook configuration, status lists, day patterns and timeslots",
        python_import=True,
    ),
    "STUDENTS": SheetSourceRule(
        sheet="STUDENTS",
        authority=SheetAuthority.AUTHORITATIVE,
        purpose="Student identity, placement dates, supervisor and capabilities",
        python_import=True,
    ),
    "SESSIONS": SheetSourceRule(
        sheet="SESSIONS",
        authority=SheetAuthority.AUDIT_ONLY,
        purpose="Legacy session table; audited but excluded from scheduling import",
    ),
    "MASTER_SCHEDULE": SheetSourceRule(
        sheet="MASTER_SCHEDULE",
        authority=SheetAuthority.DERIVED_VIEW,
        purpose="Human-readable master schedule view",
        future_write_target=True,
    ),
    "THERAPIST_DAILY": SheetSourceRule(
        sheet="THERAPIST_DAILY",
        authority=SheetAuthority.DERIVED_VIEW,
        purpose="Human-readable daily therapist view",
        future_write_target=True,
    ),
    "CONFLICT_LOG": SheetSourceRule(
        sheet="CONFLICT_LOG",
        authority=SheetAuthority.DERIVED_VIEW,
        purpose="Conflict output/log view",
        future_write_target=True,
    ),
    "REPLACEMENTS": SheetSourceRule(
        sheet="REPLACEMENTS",
        authority=SheetAuthority.DERIVED_VIEW,
        purpose="Replacement suggestions and selection UI",
        future_write_target=True,
    ),
    "REPLACEMENT_LOG": SheetSourceRule(
        sheet="REPLACEMENT_LOG",
        authority=SheetAuthority.LEGACY_OPERATIONAL,
        purpose="Historical replacement log; final Python ownership not yet assigned",
    ),
    "THERAPIST_ATTENDANCE": SheetSourceRule(
        sheet="THERAPIST_ATTENDANCE",
        authority=SheetAuthority.LEGACY_OPERATIONAL,
        purpose="Existing therapist attendance/absence workflow; redesign pending",
    ),
    "DAILY_LOG": SheetSourceRule(
        sheet="DAILY_LOG",
        authority=SheetAuthority.LEGACY_OPERATIONAL,
        purpose="Existing daily operational log; redesign pending",
    ),
    "STATISTICS": SheetSourceRule(
        sheet="STATISTICS",
        authority=SheetAuthority.DERIVED_VIEW,
        purpose="Statistics/reporting output",
        future_write_target=True,
    ),
    "NEW_PATIENT": SheetSourceRule(
        sheet="NEW_PATIENT",
        authority=SheetAuthority.UI_ONLY,
        purpose="Existing data-entry UI; PATIENTS remains the patient source of truth",
    ),
}


AUTHORITATIVE_SHEETS = frozenset(
    name
    for name, rule in SOURCE_RULES.items()
    if rule.authority is SheetAuthority.AUTHORITATIVE
)


def get_source_rule(sheet: str) -> SheetSourceRule | None:
    return SOURCE_RULES.get(sheet)


def is_authoritative(sheet: str) -> bool:
    rule = get_source_rule(sheet)
    return rule is not None and rule.authority is SheetAuthority.AUTHORITATIVE


def imported_sheets() -> frozenset[str]:
    return frozenset(name for name, rule in SOURCE_RULES.items() if rule.python_import)
