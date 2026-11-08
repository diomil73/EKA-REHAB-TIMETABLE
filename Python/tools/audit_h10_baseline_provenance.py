from __future__ import annotations

"""Read-only comparison of a finalized day in trusted baseline and demo candidate.

Never re-finalize, edit, mask, or copy patient identifiers to the output.
Confirms whether a problem predates the H8 synthetic seed.
"""

import argparse
from datetime import date
from hashlib import sha256
from pathlib import Path
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.report_therapist_history import report_therapist_history
from rehab_excel.patient_session_history_store import read_current_patient_session_history_rows
from rehab_core.daily_state import DailySessionStatus


def _sha(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def _summaries(workbook: Path, service_date: date) -> tuple[tuple, ...]:
    daily = [
        item for item in report_therapist_history(workbook)
        if item.service_date == service_date
    ]
    sessions = read_current_patient_session_history_rows(workbook)
    result = []
    for row in sorted(daily, key=lambda x: x.therapist_id.casefold()):
        delivered = [
            s for s in sessions
            if s.service_date == service_date
            and s.snapshot_id == row.snapshot_id
            and s.effective_therapist_id == row.therapist_id
            and s.status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED)
        ]
        distinct_times = {s.effective_time for s in delivered}
        result.append((
            row.snapshot_id, row.snapshot_version, row.therapist_id,
            row.capacity_timeslots, row.planned_occupied_timeslots,
            row.final_occupied_timeslots, row.planned_sessions,
            row.final_active_sessions,
            row.approved_shared_or_over_capacity_events,
            len(delivered), len(distinct_times),
            tuple(sorted((str(t) for t in distinct_times))),
        ))
    return tuple(result)


def audit(baseline: Path, candidate: Path, service_date: date) -> int:
    for path in (baseline, candidate):
        if not path.is_file() or path.suffix.casefold() != ".xlsm":
            print("SAFETY STOP: Existing XLSM required")
            return 2
    before = (_sha(baseline), _sha(candidate))
    try:
        original = _summaries(baseline, service_date)
        copied = _summaries(candidate, service_date)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    if not original:
        print(f"SAFETY STOP: Baseline has no current-final facts for {service_date}")
        return 2
    print(f"READ-ONLY HISTORICAL PROVENANCE {service_date.isoformat()}")
    for item in copied:
        sid, ver, person, capacity, initial, final, planned, delivered, shared, count, times, clocks = item
        print(
            f"{person} | snapshot={sid} v{ver} | capacity={capacity} | "
            f"initial_slots={initial} | final_slots={final} | "
            f"planned={planned} | delivered={delivered} | "
            f"approved_shared={shared} | delivered_rows={count} | "
            f"distinct_effective_times={times} | clocks={','.join(clocks)}"
        )
    if (_sha(baseline), _sha(candidate)) != before:
        print("SAFETY STOP: Workbook hash changed during read-only audit")
        return 2
    if original != copied:
        print("SAFETY STOP: Baseline and candidate finalized day differ")
        return 3
    print("BASELINE DAY IDENTICAL IN CANDIDATE; no workbook modified")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only baseline/day origin proof")
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--date", type=date.fromisoformat, default=date(2026, 10, 30))
    args = parser.parse_args()
    return audit(args.baseline, args.candidate, args.date)


if __name__ == "__main__":
    raise SystemExit(main())
