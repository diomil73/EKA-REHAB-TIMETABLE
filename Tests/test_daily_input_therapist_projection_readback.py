"""Synthetic, non-authorizing THERAPIST_DAILY occupied-clock checks."""
from datetime import time
import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.daily_input_therapist_projection_readback import (
    TherapistDailyProjectionReadbackError,
    verify_therapist_daily_occupied_clocks,
)


def fixture(tmp_path):
    path = tmp_path / "SYNTHETIC_THERAPIST_DAILY.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "THERAPIST_DAILY"
    ws["B1"] = "THERAPIST A"
    ws["C1"] = "THERAPIST B"
    ws["B11"] = "THERAPIST A"
    ws["A2"] = "08:00"
    ws["A3"] = "09:00"
    ws["A12"] = "08:00"
    ws["A13"] = "10:00"
    ws["B2"] = "FICTIONAL P1"
    ws["B3"] = "FICTIONAL P2"
    ws["C2"] = "FICTIONAL P3"
    ws["B12"] = "FICTIONAL P4"
    ws["B13"] = "FICTIONAL P5"
    wb.save(path)
    wb.close()
    expected = (
        ("THERAPIST A", (time(8), time(9), time(10))),
        ("THERAPIST B", (time(8),)),
    )
    return path, expected


def test_saved_therapist_occupied_clock_count_distinct_across_blocks(tmp_path):
    path, expected = fixture(tmp_path)
    before = path.read_bytes()
    result = verify_therapist_daily_occupied_clocks(path, expected_clocks=expected)
    assert result.checked_providers == 2
    assert result.occupied_distinct_clocks == 4
    assert path.read_bytes() == before


@pytest.mark.parametrize("change", [
    "wrong-time", "missing-slot", "extra-slot", "extra-provider",
    "header-error", "orphan-occupancy", "formula-occupancy",
    "formula-clock", "missing-sheet",
])
def test_rejects_saved_therapist_view_mismatch(tmp_path, change):
    path, expected = fixture(tmp_path)
    wb = load_workbook(path, keep_vba=False)
    ws = wb["THERAPIST_DAILY"]
    if change == "wrong-time":
        ws["A3"] = "11:00"
    elif change == "missing-slot":
        ws["B3"] = None
    elif change == "extra-slot":
        ws["A4"], ws["B4"] = "11:00", "P6"
    elif change == "extra-provider":
        ws["D1"], ws["D2"] = "THERAPIST C", "P7"
    elif change == "header-error":
        ws["B1"] = "=1+1"
    elif change == "orphan-occupancy":
        ws["D2"] = "UNKNOWN"
    elif change == "formula-occupancy":
        ws["B2"] = "=1+1"
    elif change == "formula-clock":
        ws["A2"] = "=1/3"
    else:
        ws.title = "OTHER"
    wb.save(path)
    wb.close()
    original = path.read_bytes()
    with pytest.raises(TherapistDailyProjectionReadbackError):
        verify_therapist_daily_occupied_clocks(path, expected_clocks=expected)
    assert path.read_bytes() == original


@pytest.mark.parametrize("invalid", [
    (),
    (("THERAPIST A", (time(8),)), ("therapist a", (time(9),))),
    (("THERAPIST A", (time(8), time(8))),),
    (("THERAPIST A", [time(8)]),),
])
def test_refuses_missing_duplicated_or_mutable_preflight(tmp_path, invalid):
    path, _ = fixture(tmp_path)
    with pytest.raises(TherapistDailyProjectionReadbackError):
        verify_therapist_daily_occupied_clocks(path, expected_clocks=invalid)


def test_no_default_capacity_or_journal_permission():
    import inspect
    import rehab_excel.daily_input_therapist_projection_readback as module
    source = inspect.getsource(module)
    assert "DEFAULT_CAPACITY" not in source
    assert "DailyInputJournalApplyEnabled = True" not in source
    assert "DispatchEx" not in source


def test_clock_rounding_must_not_wrap_24h_to_midnight(tmp_path):
    path, expected = fixture(tmp_path)
    wb = load_workbook(path, keep_vba=False)
    # A saved numeric time can be just below 24:00 and round to 1440
    # minutes. This must never silently masquerade as 00:00.
    wb["THERAPIST_DAILY"]["A2"] = 0.9999999
    wb.save(path)
    wb.close()
    with pytest.raises(TherapistDailyProjectionReadbackError, match="day boundary"):
        verify_therapist_daily_occupied_clocks(path, expected_clocks=expected)


@pytest.mark.parametrize("clock", [time(8, 0, 1), time(8, 0, 0, 1)])
def test_expected_clock_requires_exact_minute_resolution(tmp_path, clock):
    path, expected = fixture(tmp_path)
    changed = ((expected[0][0], (clock, time(9), time(10))), expected[1])
    with pytest.raises(TherapistDailyProjectionReadbackError, match="Malformed"):
        verify_therapist_daily_occupied_clocks(path, expected_clocks=changed)
