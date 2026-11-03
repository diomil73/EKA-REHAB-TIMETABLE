from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

from rehab_core.daily_state import DailySessionStatus
from rehab_excel.patient_session_history_store import (
    read_current_patient_session_history_rows,
)


class PatientHistoryReportError(ValueError):
    """Raised if finalized patient history is unsuitable for presentation."""


@dataclass(frozen=True)
class PatientHistoryReportRow:
    period_start: date
    period_end: date
    patient_id: str
    finalized_days: int
    scheduled: int
    performed: int
    patient_absences: int
    cancellations: int
    replacements: int


def _period_range(day: date, mode: str) -> tuple[date, date]:
    if mode == "daily":
        return day, day
    if mode == "weekly":
        start = day - timedelta(days=day.weekday())
        return start, start + timedelta(days=6)
    if mode == "monthly":
        first = day.replace(day=1)
        if day.month == 12:
            next_month = date(day.year + 1, 1, 1)
        else:
            next_month = date(day.year, day.month + 1, 1)
        return first, next_month - timedelta(days=1)
    if mode == "overall":
        return date.min, date.max
    raise PatientHistoryReportError(f"Unsupported report mode: {mode!r}")


def _patient_sort_key(value: str) -> tuple[int, int | str, str]:
    """Natural order for numeric patient identifiers without altering IDs."""
    if value.isascii() and value.isdecimal():
        return (0, int(value), value)
    return (1, value.casefold(), value)


def summarize_patient_history_report(
    rows: tuple,
    *,
    mode: str,
    start: date | None = None,
    end: date | None = None,
) -> tuple[PatientHistoryReportRow, ...]:
    """Summarize already-version-filtered patient session history.

    Callers must use read_current_patient_session_history_rows, not the full
    immutable audit history. Duplicate finalized session keys stop the report.
    """
    mode = mode.strip().casefold()
    if mode not in {"daily", "weekly", "monthly", "overall"}:
        raise PatientHistoryReportError(f"Unsupported report mode: {mode!r}")
    if start is not None and end is not None and start > end:
        raise PatientHistoryReportError("start date after end date")

    groups: dict[tuple[date, date, str], dict[str, object]] = {}
    seen: set[tuple[date, str]] = set()

    for row in rows:
        if start is not None and row.service_date < start:
            continue
        if end is not None and row.service_date > end:
            continue
        if row.status == DailySessionStatus.THERAPIST_ABSENT:
            raise PatientHistoryReportError("Unresolved session in finalized history")
        patient_id = row.patient_id.strip()
        if not patient_id:
            raise PatientHistoryReportError("Blank patient identifier")
        session_key = (row.service_date, row.session_id.casefold())
        if session_key in seen:
            raise PatientHistoryReportError("Duplicate CURRENT_FINAL session")
        seen.add(session_key)

        beginning, ending = _period_range(row.service_date, mode)
        key = (beginning, ending, patient_id)
        if key not in groups:
            groups[key] = {
                "dates": set(), "scheduled": 0, "performed": 0,
                "patient_absences": 0, "cancellations": 0, "replacements": 0,
                "first": row.service_date, "last": row.service_date,
            }
        item = groups[key]
        item["dates"].add(row.service_date)
        item["first"] = min(item["first"], row.service_date)
        item["last"] = max(item["last"], row.service_date)
        item["scheduled"] += 1
        if row.status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED):
            item["performed"] += 1
        if row.status == DailySessionStatus.PATIENT_ABSENT:
            item["patient_absences"] += 1
        if row.status == DailySessionStatus.CANCELLED:
            item["cancellations"] += 1
        if row.status == DailySessionStatus.REPLACED:
            item["replacements"] += 1

    result = []
    for (beginning, ending, patient_id), item in sorted(
        groups.items(), key=lambda pair: (pair[0][0], _patient_sort_key(pair[0][2]))
    ):
        # For overall, report actual observed date span, not date.min/date.max.
        if mode == "overall":
            beginning, ending = item["first"], item["last"]
        result.append(PatientHistoryReportRow(
            period_start=beginning,
            period_end=ending,
            patient_id=patient_id,
            finalized_days=len(item["dates"]),
            scheduled=item["scheduled"],
            performed=item["performed"],
            patient_absences=item["patient_absences"],
            cancellations=item["cancellations"],
            replacements=item["replacements"],
        ))
    return tuple(result)


def read_patient_history_report(
    workbook_path: str | Path,
    *,
    mode: str = "monthly",
    start: date | None = None,
    end: date | None = None,
) -> tuple[PatientHistoryReportRow, ...]:
    """Read-only patient report from the authoritative CURRENT_FINAL H5 reader."""
    return summarize_patient_history_report(
        read_current_patient_session_history_rows(workbook_path),
        mode=mode, start=start, end=end,
    )


def include_registered_patients_without_sessions(
    reports: tuple[PatientHistoryReportRow, ...],
    registry_ids: tuple[str, ...],
    *,
    mode: str,
) -> tuple[PatientHistoryReportRow, ...]:
    """Explicitly include current-registry members without finalized activity.

    These zero rows are presentation-only, not historical records. This is a
    present-day registry overlay and cannot reconstruct prior registry rosters.
    """
    mode = mode.strip().casefold()
    if mode not in {"daily", "weekly", "monthly", "overall"}:
        raise PatientHistoryReportError(f"Unsupported report mode: {mode!r}")

    def identity(value: str) -> tuple[str, str | int]:
        clean = str(value).strip()
        if not clean:
            raise PatientHistoryReportError("Blank registry patient identifier")
        if clean.isascii() and clean.isdecimal():
            return ("numeric", int(clean))
        return ("text", clean.casefold())

    registry: dict[tuple[str, str | int], str] = {}
    for original in registry_ids:
        key = identity(original)
        if key in registry:
            raise PatientHistoryReportError("Duplicate patient ID in registry")
        registry[key] = str(original).strip()

    if not reports:
        return reports  # No finalized period: never invent a reporting period.

    if mode == "overall":
        overall_period = (
            min(row.period_start for row in reports),
            max(row.period_end for row in reports),
        )
        periods = {overall_period}
    else:
        periods = {(row.period_start, row.period_end) for row in reports}

    existing: dict[tuple[date, date], set[tuple[str, str | int]]] = {
        period: set() for period in periods
    }
    for row in reports:
        patient_key = identity(row.patient_id)
        if patient_key not in registry:
            raise PatientHistoryReportError(
                "CURRENT_FINAL patient is absent from current registry"
            )
        period = overall_period if mode == "overall" else (
            row.period_start, row.period_end
        )
        if patient_key in existing[period]:
            raise PatientHistoryReportError("Duplicate patient period row")
        existing[period].add(patient_key)

    extended = list(reports)
    for beginning, ending in sorted(periods):
        for key, patient_id in registry.items():
            if key in existing[(beginning, ending)]:
                continue
            extended.append(PatientHistoryReportRow(
                period_start=beginning,
                period_end=ending,
                patient_id=patient_id,
                finalized_days=0,
                scheduled=0,
                performed=0,
                patient_absences=0,
                cancellations=0,
                replacements=0,
            ))
    return tuple(sorted(
        extended,
        key=lambda row: (
            row.period_start if mode != "overall" else date.min,
            row.period_end if mode != "overall" else date.min,
            _patient_sort_key(row.patient_id),
        ),
    ))
