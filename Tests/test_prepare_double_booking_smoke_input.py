from datetime import date, time
from pathlib import Path
import importlib.util

from rehab_core.models import Session


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "prepare_double_booking_smoke_input.py"


def _load_tool_module():
    spec = importlib.util.spec_from_file_location(
        "prepare_double_booking_smoke_input", TOOL_PATH
    )
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _session(
    session_id: str,
    patient_id: str,
    therapist_id: str,
    hh: int,
    minute: int = 0,
):
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


def test_choose_double_booking_smoke_case_skips_provider_with_two_sessions_same_slot():
    module = _load_tool_module()
    target, _busy = module.choose_double_booking_smoke_case(
        (
            _session("A1", "P1", "T1", 10),
            _session("A2", "P2", "T1", 10),
            _session("B", "P3", "T2", 10),
            _session("C", "P4", "T3", 11),
            _session("D", "P5", "T4", 11),
        ),
        therapist_names=("T1", "T2", "T3", "T4"),
    )

    assert target.therapist_id != "T1"


def test_blocked_therapists_blocks_free_registry_providers_plus_original():
    module = _load_tool_module()
    target = _session("A", "P1", "T1", 10)
    sessions = (
        target,
        _session("B", "P2", "T2", 10),
        _session("C", "P3", "T3", 11),
    )

    blocked = module.blocked_therapists_for_smoke(
        sessions,
        target=target,
        therapist_names=("T1", "T2", "T3", "T4"),
    )

    assert blocked == ("T1", "T3", "T4")
    assert "T2" not in blocked


def test_smoke_preparer_uses_fresh_daily_input_copy_and_application_contract():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "create_daily_input_preview(" in source
    assert "assert_application_contract(source)" in source
    assert "assert_application_contract(output)" in source
    assert "file_sha256(source)" in source
    assert 'ws.Cells(row, 4).Value = "ΜΗ ΔΙΑΘΕΣΙΜΟΣ"' in source


def test_smoke_preparer_uses_one_slot_absences_not_whole_day_absences():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert 'ws.Cells(row, 2).Value = "ΟΧΙ"' in source
    assert 'ws.Cells(row, 3).Value = original_time.strftime("%H:%M")' in source
    assert 'ws.Cells(row, 4).Value = original_time.strftime("%H:%M")' in source
    assert "Block normal replacement only in target slot" in source
    assert 'ws.Cells(row, 2).Value = "ΝΑΙ"' not in source


def test_smoke_preparer_expands_disposable_therapist_section_when_needed():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "def _ensure_therapist_row_capacity(" in source
    assert 'ws.Rows(f"{insert_start}:{insert_end}").Insert()' in source
    assert "patient_header = _ensure_therapist_row_capacity(" in source
    assert "insufficient therapist rows for smoke" not in source


def test_smoke_preparer_uses_full_registry_not_only_materialized_providers():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "therapist_names = tuple(settings.therapist_names)" in source
    assert "therapist_names=therapist_names" in source
    assert "normal_registry = {" in source
    assert "if not _is_legacy_student_name(therapist_id)" in source
    assert "active_day_therapists" not in source
