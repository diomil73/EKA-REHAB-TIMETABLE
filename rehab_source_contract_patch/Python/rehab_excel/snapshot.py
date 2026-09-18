from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from rehab_core.models import BaseScheduleEntry, Patient

from .reader import WorkbookAudit, WorkbookSettings, audit_workbook, read_base_schedule, read_patients, read_settings
from .source_contract import AUTHORITATIVE_SHEETS


class SourceContractError(RuntimeError):
    """Raised when a workbook violates an authoritative-source safety rule."""


@dataclass(frozen=True)
class WorkbookSnapshot:
    """Immutable in-memory view of the confirmed workbook sources.

    Building a snapshot is read-only. It does not persist patient data outside
    the workbook and it does not save or modify the .xlsm.
    """

    workbook_path: str
    patients: tuple[Patient, ...]
    settings: WorkbookSettings
    base_schedule: tuple[BaseScheduleEntry, ...]
    audit: WorkbookAudit

    @property
    def patient_count(self) -> int:
        return len(self.patients)

    @property
    def base_entry_count(self) -> int:
        return len(self.base_schedule)

    @property
    def authoritative_sheets(self) -> frozenset[str]:
        return AUTHORITATIVE_SHEETS


def build_snapshot(path: str | Path, *, require_clean_contract: bool = True) -> WorkbookSnapshot:
    """Build a read-only snapshot from authoritative workbook sources.

    Warnings such as unresolved blank day patterns do not block the snapshot.
    Contract errors, such as duplicate PatientIDs or PATIENTS/PLANNER identity
    disagreement, block it when ``require_clean_contract`` is True.
    """

    audit = audit_workbook(path)
    if require_clean_contract and not audit.ok:
        errors = [
            f"{issue.code} ({issue.count})"
            for issue in audit.issues
            if issue.severity == "error"
        ]
        raise SourceContractError(
            "Workbook violates the authoritative source contract: " + ", ".join(errors)
        )

    return WorkbookSnapshot(
        workbook_path=str(path),
        patients=tuple(read_patients(path)),
        settings=read_settings(path),
        base_schedule=tuple(read_base_schedule(path)),
        audit=audit,
    )
