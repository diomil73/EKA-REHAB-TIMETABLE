from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook

from tools import inspect_close_day_native_outcome as module
from rehab_excel.day_history_event_store import HEADERS as EVENT_HEADERS
from rehab_excel.finalized_day_snapshots import HEADERS as SNAPSHOT_HEADERS
from rehab_excel.therapist_day_history_store import HEADERS as THERAPIST_HEADERS


def _write_row(ws, row: int, values):
    for col, value in enumerate(values, start=1):
        ws.cell(row, col).value = value


def test_inspector_accepts_one_consistent_finalized_day(tmp_path: Path):
    path = tmp_path / "close_day_smoke.xlsm"
    wb = Workbook()
    daily = wb.active
    daily.title = "DAILY_INPUT"
    daily["B2"] = "30/10/2026"

    finalized_at = datetime(2026, 11, 1, 19, 0, tzinfo=timezone.utc).isoformat()
    snapshot_id = "finalized:2026-10-30:v1"

    snapshots = wb.create_sheet("FINALIZED_DAY_SNAPSHOTS")
    _write_row(snapshots, 1, SNAPSHOT_HEADERS)
    _write_row(
        snapshots,
        2,
        (
            snapshot_id,
            "2026-10-30",
            1,
            "CURRENT_FINAL",
            finalized_at,
            "",
            0,
            10,
            9,
            1,
            0,
            0,
            0,
            0,
            0,
        ),
    )
    snapshots.sheet_state = "veryHidden"

    therapists = wb.create_sheet("THERAPIST_DAY_HISTORY")
    _write_row(therapists, 1, THERAPIST_HEADERS)
    _write_row(
        therapists,
        2,
        (snapshot_id, "2026-10-30", "THERAPIST A", 6, 5, 5, 4, 4, 0, 1, 0, 0),
    )
    therapists.sheet_state = "veryHidden"

    events = wb.create_sheet("DAY_HISTORY_EVENTS")
    _write_row(events, 1, EVENT_HEADERS)
    _write_row(
        events,
        2,
        (
            "day-finalized:2026-10-30:v1",
            "2026-10-30",
            "DAY_FINALIZED",
            finalized_at,
            "close-day-action",
            "",
            "",
            "",
            "",
            "",
            snapshot_id,
            "Day finalized",
        ),
    )
    events.sheet_state = "veryHidden"
    wb.save(path)

    result = module.inspect_close_day_outcome(path)

    assert result["service_date"] == "2026-10-30"
    assert result["snapshot_id"] == snapshot_id
    assert result["snapshot_status"] == "CURRENT_FINAL"
    assert result["therapist_history_rows"] == 1
    assert result["event_id"] == "day-finalized:2026-10-30:v1"
    assert set(result["history_sheet_states"].values()) == {"veryHidden"}


def test_inspector_rejects_missing_finalized_snapshot(tmp_path: Path):
    path = tmp_path / "not_closed.xlsm"
    wb = Workbook()
    daily = wb.active
    daily.title = "DAILY_INPUT"
    daily["B2"] = "30/10/2026"
    wb.save(path)

    try:
        module.inspect_close_day_outcome(path)
    except module.CloseDayInspectionError as exc:
        assert "Expected exactly one CURRENT_FINAL snapshot" in str(exc)
    else:
        raise AssertionError("Inspector accepted a workbook with no finalized snapshot")
