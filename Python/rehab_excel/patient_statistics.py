from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from rehab_core.daily_state import DailySessionStatus
from rehab_excel.patient_session_history_store import (
    read_current_patient_session_history_rows,
)


@dataclass(frozen=True)
class PatientStatistics:
    patient_id: str
    scheduled: int
    performed: int
    absent: int
    cancelled: int
    replaced: int


def summarize_patient_sessions(
    rows: tuple,
    *,
    start: date | None = None,
    end: date | None = None,
) -> tuple[PatientStatistics, ...]:
    """Summarize only supplied CURRENT_FINAL rows, never historical versions.

    ACTIVE and REPLACED are delivered sessions. The cancellation-kind breakdown
    (including no-shows) will be a separately specified metric.
    """
    if start is not None and end is not None and start > end:
        raise ValueError("start date cannot be after end date")

    totals: dict[str, Counter[str]] = {}
    seen: set[tuple[date, str]] = set()
    for row in rows:
        if start is not None and row.service_date < start:
            continue
        if end is not None and row.service_date > end:
            continue
        key = (row.service_date, row.session_id.casefold())
        if key in seen:
            raise ValueError(f"Duplicate CURRENT_FINAL session {key!r}")
        seen.add(key)
        if row.status == DailySessionStatus.THERAPIST_ABSENT:
            raise ValueError("Unresolved therapist absence in finalized statistics")
        patient = row.patient_id.strip()
        if not patient:
            raise ValueError("Blank patient id")
        counts = totals.setdefault(patient, Counter())
        counts["scheduled"] += 1
        if row.status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED):
            counts["performed"] += 1
        if row.status == DailySessionStatus.PATIENT_ABSENT:
            counts["absent"] += 1
        if row.status == DailySessionStatus.CANCELLED:
            counts["cancelled"] += 1
        if row.status == DailySessionStatus.REPLACED:
            counts["replaced"] += 1

    return tuple(
        PatientStatistics(
            patient_id=patient,
            scheduled=counts["scheduled"],
            performed=counts["performed"],
            absent=counts["absent"],
            cancelled=counts["cancelled"],
            replaced=counts["replaced"],
        )
        for patient, counts in sorted(totals.items(), key=lambda item: item[0].casefold())
    )


def read_patient_statistics(
    workbook_path: str | Path,
    *,
    start: date | None = None,
    end: date | None = None,
) -> tuple[PatientStatistics, ...]:
    """Read-only stats bridge: existing version-aware H5 reader is authoritative."""
    return summarize_patient_sessions(
        read_current_patient_session_history_rows(workbook_path),
        start=start,
        end=end,
    )
