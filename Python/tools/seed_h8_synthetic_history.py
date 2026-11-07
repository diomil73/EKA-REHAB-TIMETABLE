from __future__ import annotations

"""One-shot native Excel H8 demo-history seeder for a fresh, disposable XLSM.

NEVER invoke on BL/BK/BJ/AZ or on a production workbook. Input is a newly
created candidate from prepare_h8_dashboard_smoke.py; this program appends
synthetic CURRENT_FINAL history while preserving VBA using Excel COM.
"""

import argparse
from collections import Counter
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
import hashlib
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.finalized_day_snapshots import (
    HEADERS as SNAP_HEADERS, read_finalized_day_snapshots,
)
from rehab_excel.therapist_day_history_store import (
    HEADERS as THERAPIST_HEADERS, read_therapist_day_history_rows,
)
from rehab_excel.patient_session_history_store import (
    HEADERS as PATIENT_HEADERS, read_current_patient_session_history_rows,
)
from rehab_excel.day_history_event_store import HEADERS as EVENT_HEADERS
from rehab_excel.therapist_attendance_history_store import (
    HEADERS as ATTENDANCE_HEADERS, SHEET_NAME as ATTENDANCE_SHEET,
    read_verified_therapist_attendance,
)
from rehab_excel.patient_registry_source import read_patient_registry
from rehab_core.daily_state import DailySessionStatus
from rehab_core.therapist_history_reports import build_therapist_daily_reports


HISTORY_SHEETS = (
    ("FINALIZED_DAY_SNAPSHOTS", SNAP_HEADERS),
    ("THERAPIST_DAY_HISTORY", THERAPIST_HEADERS),
    ("PATIENT_SESSION_HISTORY", PATIENT_HEADERS),
    ("DAY_HISTORY_EVENTS", EVENT_HEADERS),
)
DEMO_SOURCE = "SYNTHETIC_DEMO_APPROVED_ROSTER_NOT_REAL"
FROZEN_SUFFIXES = ("_AZ.xlsm", "_BJ.xlsm", "_BK.xlsm", "_BL.xlsm")


class DemoHistoryError(RuntimeError):
    pass


def _days(begin: date, end: date, excluded: set[date]) -> tuple[date, ...]:
    result = []
    current = begin
    while current <= end:
        if current.weekday() < 5 and current not in excluded:
            result.append(current)
        current += timedelta(days=1)
    return tuple(result)


def _patient_id(value: object) -> str:
    return str(value).strip()


def generate_demo_rows(
    days: tuple[date, ...], therapists: tuple[str, ...],
    patients: tuple[str, ...],
) -> dict[str, tuple[tuple[object, ...], ...]]:
    """Deterministic, internally consistent current-final synthetic history."""
    if len(days) < 12 or len(set(days)) != len(days):
        raise DemoHistoryError("At least 12 unique demo dates required")
    if len(therapists) < 6 or len({x.casefold() for x in therapists}) != len(therapists):
        raise DemoHistoryError("Six distinct existing therapist IDs required")
    if len(patients) < 24 or len({x.casefold() for x in patients}) != len(patients):
        raise DemoHistoryError("24 distinct existing patient IDs required")
    if any(day.weekday() >= 5 for day in days):
        raise DemoHistoryError("Synthetic operational history must exclude weekends")
    if any(not x.strip() for x in therapists + patients):
        raise DemoHistoryError("Blank roster identifier")

    t = therapists[:6]
    outputs: dict[str, list[tuple[object, ...]]] = {
        key: [] for key in ("FINALIZED_DAY_SNAPSHOTS", "THERAPIST_DAY_HISTORY",
                            "PATIENT_SESSION_HISTORY", "DAY_HISTORY_EVENTS",
                            ATTENDANCE_SHEET)
    }
    for i, day in enumerate(days):
        snapshot = f"DEMO-H8-{day:%Y%m%d}-V1"
        finalized_at = f"{day.isoformat()}T17:30:00+02:00"
        absent_idx = i % 6 if i % 5 == 0 else None
        # One possible completely absent therapist/day. Every patient
        # scheduled with that therapist is explicitly reassigned.
        statuses = Counter()
        planned_by = Counter()
        actual_by = Counter()
        planned_sessions = Counter()
        actual_sessions = Counter()
        replacement_by = Counter()
        released_by = Counter()
        suspension_by = Counter()
        planned_slots: dict[str, set[str]] = {x: set() for x in t}
        actual_slots: dict[str, set[str]] = {x: set() for x in t}
        patient_seen: set[str] = set()
        session_rows = []
        for j, original in enumerate(t):
            scheduled = 1 + ((i * 3 + j * 2) % 6)
            for k in range(scheduled):
                patient = patients[(i * 13 + j * 6 + k) % len(patients)]
                # A patient can have more than one type of treatment in
                # reality. This demo keeps one patient/day for a clear chart.
                if patient in patient_seen:
                    patient = next((p for p in patients if p not in patient_seen), "")
                if not patient:
                    raise DemoHistoryError("Insufficient distinct patients for demo day")
                patient_seen.add(patient)
                original_time = f"{8 + k:02d}:00:00"
                effective_therapist = original
                effective_time = original_time
                replacement_id = ""
                cancellation_id = ""
                cancellation_kind = ""
                reason = ""
                pattern = (i * 17 + j * 11 + k * 7) % 19
                if j == absent_idx:
                    status = DailySessionStatus.REPLACED
                    effective_therapist = t[(j + 1) % 6]
                    effective_time = f"{14 + k:02d}:00:00"
                    replacement_id = f"DEMO-REPL-{day:%Y%m%d}-{j}-{k}"
                    reason = "Synthetic therapist sick leave replacement"
                elif pattern in (0, 1):
                    status = DailySessionStatus.PATIENT_ABSENT
                    effective_therapist = ""
                    effective_time = ""
                    reason = "Synthetic patient absence"
                elif pattern == 2:
                    status = DailySessionStatus.CANCELLED
                    effective_therapist = ""
                    effective_time = ""
                    cancellation_id = f"DEMO-CANCEL-{day:%Y%m%d}-{j}-{k}"
                    cancellation_kind = "department_postponed"
                    reason = "Synthetic department postponement"
                elif pattern == 3:
                    status = DailySessionStatus.CANCELLED
                    effective_therapist = ""
                    effective_time = ""
                    cancellation_id = f"DEMO-SUSP-{day:%Y%m%d}-{j}-{k}"
                    cancellation_kind = "treatment_suspended"
                    reason = "Synthetic treatment suspension"
                else:
                    status = DailySessionStatus.ACTIVE
                statuses[status.value] += 1
                planned_by[original] += 1
                planned_sessions[original] += 1
                planned_slots[original].add(original_time)
                if status in (DailySessionStatus.ACTIVE, DailySessionStatus.REPLACED):
                    actual_by[effective_therapist] += 1
                    actual_sessions[effective_therapist] += 1
                    actual_slots[effective_therapist].add(effective_time)
                    if status == DailySessionStatus.REPLACED:
                        replacement_by[effective_therapist] += 1
                elif status == DailySessionStatus.PATIENT_ABSENT or cancellation_kind == "department_postponed":
                    released_by[original] += 1
                if cancellation_kind == "treatment_suspended":
                    suspension_by[original] += 1
                session_rows.append((
                    snapshot, day.isoformat(), 1,
                    f"DEMO-SESSION-{day:%Y%m%d}-{j}-{k}", patient,
                    status.value, original, original_time,
                    effective_therapist, effective_time, replacement_id,
                    cancellation_id, cancellation_kind, reason,
                ))
        if not session_rows:
            raise DemoHistoryError("Cannot create empty finalized demo day")
        for therapist in t:
            if len(actual_slots[therapist]) != actual_sessions[therapist]:
                raise DemoHistoryError("Overlapping therapist times in synthetic data")
            outputs["THERAPIST_DAY_HISTORY"].append((
                snapshot, day.isoformat(), therapist, 6,
                len(planned_slots[therapist]), planned_sessions[therapist],
                len(actual_slots[therapist]), actual_sessions[therapist],
                replacement_by[therapist], released_by[therapist],
                suspension_by[therapist], 0,
            ))
            absent = therapist == t[absent_idx] if absent_idx is not None else False
            if absent and actual_sessions[therapist]:
                raise DemoHistoryError("Absent demo therapist has delivered sessions")
            outputs[ATTENDANCE_SHEET].append((
                snapshot, day.isoformat(), therapist,
                "ABSENT" if absent else "PRESENT", DEMO_SOURCE,
                "SICK_LEAVE" if absent else "",
            ))
        outputs["FINALIZED_DAY_SNAPSHOTS"].append((
            snapshot, day.isoformat(), 1, "CURRENT_FINAL",
            finalized_at, "", 0, len(session_rows),
            statuses["active"] + statuses["replaced"],
            statuses["patient_absent"], statuses["cancelled"],
            0, statuses["replaced"],
            sum(suspension_by.values()), 0,
        ))
        outputs["PATIENT_SESSION_HISTORY"].extend(session_rows)
        outputs["DAY_HISTORY_EVENTS"].append((
            f"DEMO-EVENT-{day:%Y%m%d}", day.isoformat(),
            "DAY_FINALIZED", finalized_at, DEMO_SOURCE,
            "", "", "", "", "", snapshot,
            "Synthetic finalized day for reporting demonstration only",
        ))
    return {name: tuple(records) for name, records in outputs.items()}


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _load_roster(path: Path):
    existing = read_finalized_day_snapshots(path)
    history = read_therapist_day_history_rows(path)
    therapist_ids = tuple(sorted({r.therapist_id for r in history}, key=str.casefold))
    patient_ids = tuple(_patient_id(p.patient_id) for p in read_patient_registry(path))
    if len(therapist_ids) < 6:
        raise DemoHistoryError("Baseline has fewer than six historical therapist IDs")
    if len(patient_ids) < 24:
        raise DemoHistoryError("Baseline has fewer than 24 registered patients")
    existing_dates = {x.service_date for x in existing}
    dates = _days(date(2026, 9, 21), date(2026, 10, 23), existing_dates)
    if len(dates) < 12:
        raise DemoHistoryError("Not enough non-overlapping synthetic service dates")
    return generate_demo_rows(dates, therapist_ids[:6], patient_ids), dates, existing


def seed_native(path: Path) -> None:
    """May write ONLY an explicitly named fresh demo workbook once."""
    if sys.platform != "win32":
        raise DemoHistoryError("Native Excel seed requires Windows")
    if path.suffix.casefold() != ".xlsm" or not path.is_file():
        raise DemoHistoryError("Existing fresh XLSM candidate required")
    if path.name.upper().endswith(FROZEN_SUFFIXES) or "_DEMO_" not in path.name.upper():
        raise DemoHistoryError("Refusing non-demo or frozen workbook path")
    rows, days, original_snapshots = _load_roster(path)
    if any(x.snapshot_id.startswith("DEMO-") for x in original_snapshots):
        raise DemoHistoryError("Demo already seeded. Never retry in same file")
    input_hash = _sha(path)
    import win32com.client  # type: ignore[import-not-found]
    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        book = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        if bool(book.ReadOnly):
            raise DemoHistoryError("Native destination opened read-only")
        for name, headers in (*HISTORY_SHEETS, (ATTENDANCE_SHEET, ATTENDANCE_HEADERS)):
            try:
                ws = book.Worksheets(name)
                if name == ATTENDANCE_SHEET:
                    raise DemoHistoryError("Attendance demo sheet already exists")
                existing_headers = tuple(str(ws.Cells(1, col).Value or "").strip()
                                         for col in range(1, len(headers) + 1))
                if existing_headers != headers:
                    raise DemoHistoryError(f"Historical schema mismatch: {name}")
            except DemoHistoryError:
                raise
            except Exception:
                if name != ATTENDANCE_SHEET:
                    raise DemoHistoryError(f"Required history sheet missing: {name}")
        # All preflight checks completed. No overwrite of existing rows.
        for name, headers in (*HISTORY_SHEETS, (ATTENDANCE_SHEET, ATTENDANCE_HEADERS)):
            if name == ATTENDANCE_SHEET:
                ws = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
                ws.Name = name
                for col, header in enumerate(headers, 1):
                    ws.Cells(1, col).Value2 = header
            else:
                ws = book.Worksheets(name)
                ws.Unprotect("")
            start = max(2, int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row) + 1)
            if start == 2 and name != ATTENDANCE_SHEET:
                raise DemoHistoryError(f"Unexpected empty historical sheet: {name}")
            # A single 2-D native Excel write avoids thousands of per-cell
            # COM calls while preserving the existing rows and header.
            values = rows[name]
            if values:
                ws.Cells(start, 1).Resize(len(values), len(headers)).Value = tuple(
                    tuple(record) for record in values
                )
            ws.Visible = 2
            ws.Protect("", True, True, True)
        book.Save()
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()

    # Independent read-only reload. Never repair the destination if it fails.
    current = read_finalized_day_snapshots(path, current_only=True)
    after_days = {s.service_date for s in current}
    original_current_count = sum(s.status.value == "CURRENT_FINAL" for s in original_snapshots)
    if not set(days).issubset(after_days) or len(current) != original_current_count + len(days):
        raise DemoHistoryError("Persisted demo snapshot history mismatch")
    p_rows = read_current_patient_session_history_rows(path)
    t_rows = read_therapist_day_history_rows(path)
    expected_p = len(rows["PATIENT_SESSION_HISTORY"])
    expected_t = len(rows["THERAPIST_DAY_HISTORY"])
    if sum(r.snapshot_id.startswith("DEMO-") for r in p_rows) != expected_p:
        raise DemoHistoryError("Persisted demo patient-session history mismatch")
    if sum(r.snapshot_id.startswith("DEMO-") for r in t_rows) != expected_t:
        raise DemoHistoryError("Persisted demo therapist-day history mismatch")
    for therapist in sorted({r[2] for r in rows[ATTENDANCE_SHEET]}):
        summary = read_verified_therapist_attendance(
            path, therapist_id=therapist, expected_working_dates=days)
        if summary.working_days != len(days) or summary.present_days + summary.absent_days != len(days):
            raise DemoHistoryError("Persisted demo staff attendance mismatch")
    if _sha(path) == input_hash:
        raise DemoHistoryError("Demo seed did not persist any change")
    print("H8 SYNTHETIC HISTORY SEED OK")
    print(f"Synthetic finalized days: {len(days)}")
    print(f"Synthetic patient sessions: {expected_p}")
    print(f"Synthetic therapist/day rows: {expected_t}")
    print("Synthetic attendance provenance: DEMO ONLY")
    print("Persisted current-final readback: OK")
    print("Existing history preserved; source / frozen previews untouched")


def main() -> int:
    parser = argparse.ArgumentParser(description="One-shot demo history seed (not production data)")
    parser.add_argument("--workbook", type=Path, required=True)
    args = parser.parse_args()
    try:
        seed_native(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        print("Freeze candidate; never repeat or repair seed in same XLSM.")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
