from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.therapist_history_reports import build_therapist_daily_reports  # noqa: E402
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots  # noqa: E402
from rehab_excel.therapist_day_history_store import read_therapist_day_history_rows  # noqa: E402


def _parse_date(value: str | None) -> date | None:
    if not value:
        return None
    return date.fromisoformat(value)


def _pct(value: float | None) -> str:
    if value is None:
        return "n/a"
    return f"{value * 100:.1f}%"


def report_therapist_history(
    workbook_path: str | Path,
    *,
    therapist_id: str | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
):
    path = Path(workbook_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")

    snapshots = read_finalized_day_snapshots(path)
    rows = read_therapist_day_history_rows(path)
    return build_therapist_daily_reports(
        snapshots,
        rows,
        therapist_id=therapist_id,
        start_date=start_date,
        end_date=end_date,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Read therapist historical workload from finalized workbook history.")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--therapist")
    parser.add_argument("--from-date")
    parser.add_argument("--to-date")
    args = parser.parse_args()

    try:
        reports = report_therapist_history(
            args.workbook,
            therapist_id=args.therapist,
            start_date=_parse_date(args.from_date),
            end_date=_parse_date(args.to_date),
        )
    except Exception as exc:
        print(f"THERAPIST HISTORY REPORT FAILED: {exc}")
        return 2

    print("THERAPIST HISTORY REPORT OK")
    print(f"Rows: {len(reports)}")
    for report in reports:
        print(
            " | ".join(
                (
                    report.service_date.isoformat(),
                    f"v{report.snapshot_version}",
                    report.therapist_id,
                    f"planned_slots={report.planned_occupied_timeslots}",
                    f"actual_slots={report.final_occupied_timeslots}",
                    f"planned_sessions={report.planned_sessions}",
                    f"active_sessions={report.final_active_sessions}",
                    f"replacements={report.replacements_received}",
                    f"released={report.released_by_cancellation_or_no_show}",
                    f"suspensions={report.treatment_suspensions_affecting_load}",
                    f"planned_util={_pct(report.planned_capacity_utilization)}",
                    f"actual_util={_pct(report.actual_capacity_utilization)}",
                )
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
