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
    if mode == "daily":
        rows = tuple(_daily_values(row) for row in daily)
    elif mode == "weekly":
        rows = tuple(_weekly_values(row) for row in build_therapist_weekly_reports(daily))
    elif mode == "monthly":
        rows = tuple(_period_values(row) for row in build_therapist_monthly_reports(daily))
    else:
        rows = tuple(_period_values(row) for row in build_therapist_overall_reports(daily))

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
