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
from .writeback import (
    CellPatch,
    DryRunReport,
    WriteIntent,
    WritePlan,
    WritebackSafetyError,
    build_write_plan,
    dry_run_writeback,
    validate_write_plan,
)

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
    "validate_write_plan",
    "dry_run_writeback",
    "build_write_plan",
    "WritebackSafetyError",
    "WritePlan",
    "WriteIntent",
    "DryRunReport",
    "CellPatch",
    "audit_workbook",
    "build_snapshot",
    "get_source_rule",
    "imported_sheets",
    "is_authoritative",
    "read_base_schedule",
    "read_patients",
    "read_settings",
]
