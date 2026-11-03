from datetime import date
from types import SimpleNamespace

from rehab_core.daily_state import DailySessionStatus
from tools.audit_h7_patient_report_coverage import _numeric_gap_report


def test_gap_detection_is_informational_and_numerically_sorted():
    assert _numeric_gap_report({"1", "2", "14", "16", "20"}) == tuple(
        f"{i:02d}" for i in range(3, 14)
    ) + ("15", "17", "18", "19")


def test_missing_15_is_not_invented_as_a_session():
    assert "15" in _numeric_gap_report({"13", "14", "16", "17"})
    assert "15" not in _numeric_gap_report({"13", "14", "15", "16"})


def test_non_numeric_patient_ids_are_not_assumed_missing():
    assert _numeric_gap_report({"P1", "P2"}) == ()
