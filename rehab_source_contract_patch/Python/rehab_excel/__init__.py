"""Read-only Excel integration for the Rehab Center workbook."""

from .reader import (
    AuditIssue,
    WorkbookAudit,
    WorkbookSettings,
    audit_workbook,
    read_base_schedule,
    read_patients,
    read_settings,
)
from .snapshot import SourceContractError, WorkbookSnapshot, build_snapshot
from .source_contract import (
    AUTHORITATIVE_SHEETS,
    SOURCE_RULES,
    SheetAuthority,
    SheetSourceRule,
    get_source_rule,
    imported_sheets,
    is_authoritative,
)

__all__ = [
    "AUTHORITATIVE_SHEETS",
    "AuditIssue",
    "SOURCE_RULES",
    "SheetAuthority",
    "SheetSourceRule",
    "SourceContractError",
    "WorkbookAudit",
    "WorkbookSettings",
    "WorkbookSnapshot",
    "audit_workbook",
    "build_snapshot",
    "get_source_rule",
    "imported_sheets",
    "is_authoritative",
    "read_base_schedule",
    "read_patients",
    "read_settings",
]
