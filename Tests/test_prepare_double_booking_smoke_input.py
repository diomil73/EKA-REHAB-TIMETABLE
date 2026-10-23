from datetime import date, time
from pathlib import Path
import importlib.util

from rehab_core.models import Session


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "prepare_double_booking_smoke_input.py"


def _load_tool_module():
    spec = importlib.util.spec_from_file_location("prepare_double_booking_smoke_input", TOOL_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int, minute: int = 0):
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=date(2026, 9, 18),
        start_time=time(hh, minute),
        treatment="ΦΘ",
    )


def test_choose_double_booking_smoke_case_requires_two_different_therapists_same_slot():
    module = _load_tool_module()
    target, busy = module.choose_double_booking_smoke_case(
        (
            _session("A", "P1", "T1", 10),
            _session("B", "P2", "T2", 10),
            _session("C", "P3", "T3", 11),
        ),
        therapist_names=("T1", "T2", "T3"),
    )

    assert target.start_time == busy.start_time == time(10, 0)
    assert target.therapist_id != busy.therapist_id


def test_smoke_preparer_uses_fresh_daily_input_copy_and_application_contract():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "create_daily_input_preview(" in source
    assert "assert_application_contract(source)" in source
    assert "assert_application_contract(output)" in source
    assert "file_sha256(source)" in source
    assert 'ws.Cells(row, 2).Value = "ΝΑΙ"' in source
    assert 'ws.Cells(row, 4).Value = "ΜΗ ΔΙΑΘΕΣΙΜΟΣ"' in source


def test_smoke_preparer_keeps_only_busy_exception_therapist_active():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "if therapist_id != busy_therapist" in source
    assert "if slot == original_time:" in source
    assert "Normal replacement intentionally blocked" in source
