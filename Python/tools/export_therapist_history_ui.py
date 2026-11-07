from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.therapist_history_reports import (  # noqa: E402
    build_therapist_monthly_reports,
    build_therapist_overall_reports,
    build_therapist_weekly_reports,
)
from tools.report_therapist_history import report_therapist_history  # noqa: E402
from rehab_excel.patient_session_history_store import (  # noqa: E402
    read_current_patient_session_history_rows,
)
from rehab_core.daily_state import DailySessionStatus  # noqa: E402


HEADERS = (
    "PeriodStart",
    "PeriodEnd",
    "SnapshotVersion",
    "Therapist",
    "ReportedDays",
    "PlannedSlots",
    "ActualSlots",
    "PlannedSessions",
    "ActiveSessions",
    "Replacements",
    "Released",
    "Suspensions",
    "PlannedUtilPct",
    "ActualUtilPct",
    "UniquePatientsPerDay",
)


class TherapistHistoryUiExportError(RuntimeError):
    pass


def _field(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    if "\t" in text or "\r" in text or "\n" in text:
        raise TherapistHistoryUiExportError("Export field contains unsupported control characters")
    return text


def _pct(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value * 100:.1f}"


def _daily_values(report) -> tuple[object, ...]:
    return (
        report.service_date.isoformat(),
        report.service_date.isoformat(),
        report.snapshot_version,
        report.therapist_id,
        1,
        report.planned_occupied_timeslots,
        report.final_occupied_timeslots,
        report.planned_sessions,
        report.final_active_sessions,
        report.replacements_received,
        report.released_by_cancellation_or_no_show,
        report.treatment_suspensions_affecting_load,
        _pct(report.planned_capacity_utilization),
        _pct(report.actual_capacity_utilization),
    )


def _weekly_values(report) -> tuple[object, ...]:
    return (
        report.week_start.isoformat(),
        report.week_end.isoformat(),
        "",
        report.therapist_id,
        report.reported_days,
        report.planned_occupied_timeslots,
        report.final_occupied_timeslots,
        report.planned_sessions,
        report.final_active_sessions,
        report.replacements_received,
        report.released_by_cancellation_or_no_show,
        report.treatment_suspensions_affecting_load,
        _pct(report.planned_capacity_utilization),
        _pct(report.actual_capacity_utilization),
    )


def _period_values(report) -> tuple[object, ...]:
    return (
        report.period_start.isoformat(),
        report.period_end.isoformat(),
        "",
        report.therapist_id,
        report.reported_days,
        report.planned_occupied_timeslots,
        report.final_occupied_timeslots,
        report.planned_sessions,
        report.final_active_sessions,
        report.replacements_received,
        report.released_by_cancellation_or_no_show,
        report.treatment_suspensions_affecting_load,
        _pct(report.planned_capacity_utilization),
        _pct(report.actual_capacity_utilization),
    )


def _unique_patient_days(rows: tuple) -> dict[tuple[object, str], int]:
    """COUNT DISTINCT patient IDs per therapist/day in CURRENT_FINAL sessions."""
    unique: dict[tuple[object, str], set[str]] = {}
    seen: set[tuple[object, str]] = set()
    for row in rows:
        key = (row.service_date, row.session_id.casefold())
        if key in seen:
            raise TherapistHistoryUiExportError("Duplicate CURRENT_FINAL session")
        seen.add(key)
        if row.status not in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED):
            continue
        therapist = str(row.effective_therapist_id or "").strip()
        patient = str(row.patient_id or "").strip()
        if not therapist or not patient:
            raise TherapistHistoryUiExportError("Delivered session lacks effective therapist/patient")
        unique.setdefault((row.service_date, therapist.casefold()), set()).add(patient.casefold())
    return {key: len(values) for key, values in unique.items()}


def _patient_average(report, daily_counts: dict, mode: str) -> str:
    if mode == "daily":
        start = end = report.service_date
        days = 1
    elif mode == "weekly":
        start, end = report.week_start, report.week_end
        days = report.reported_days
    else:
        start, end = report.period_start, report.period_end
        days = report.reported_days
    if days <= 0:
        return ""
    count = sum(
        patients for (day, therapist), patients in daily_counts.items()
        if start <= day <= end and therapist == report.therapist_id.casefold()
    )
    return f"{count / days:.2f}"


def export_therapist_history_ui(
    workbook_path: str | Path,
    output_path: str | Path,
    *,
    mode: str = "monthly",
) -> dict[str, object]:
    mode = mode.strip().casefold()
    if mode not in {"daily", "weekly", "monthly", "overall"}:
        raise TherapistHistoryUiExportError(f"Unsupported report mode: {mode!r}")

    daily = report_therapist_history(workbook_path)
    daily_counts = _unique_patient_days(
        read_current_patient_session_history_rows(workbook_path)
    )
    if mode == "daily":
        reports = daily
        values = _daily_values
    elif mode == "weekly":
        reports = build_therapist_weekly_reports(daily)
        values = _weekly_values
    elif mode == "monthly":
        reports = build_therapist_monthly_reports(daily)
        values = _period_values
    else:
        reports = build_therapist_overall_reports(daily)
        values = _period_values
    rows = tuple(
        (*values(report), _patient_average(report, daily_counts, mode))
        for report in reports
    )

    output = Path(output_path).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    lines = ["\t".join(HEADERS)]
    lines.extend("\t".join(_field(value) for value in row) for row in rows)
    output.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return {"ok": True, "mode": mode, "rows": len(rows), "output": str(output)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Export therapist history for the in-workbook report UI.")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--mode", choices=("daily", "weekly", "monthly", "overall"), default="monthly")
    args = parser.parse_args()

    try:
        result = export_therapist_history_ui(args.workbook, args.output, mode=args.mode)
    except Exception as exc:
        try:
            if args.output.exists():
                args.output.unlink()
        except OSError:
            pass
        print(f"THERAPIST HISTORY UI EXPORT FAILED: {exc}")
        return 2

    print("THERAPIST HISTORY UI EXPORT OK")
    print(f"Mode: {result['mode']}")
    print(f"Rows: {result['rows']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
