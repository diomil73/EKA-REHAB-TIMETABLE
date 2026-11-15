import importlib.util
from datetime import date, time
from pathlib import Path

from rehab_core.models import Session


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "prepare_global_multi_absence_smoke.py"


def _load_tool_module():
    spec = importlib.util.spec_from_file_location("prepare_global_multi_absence_smoke", TOOL_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _session(session_id: str, therapist_id: str, hour: int, minute: int = 0):
    return Session(
        session_id=session_id,
        patient_id=f"P-{session_id}",
        therapist_id=therapist_id,
        session_date=date(2026, 9, 18),
        start_time=time(hour, minute),
    )


def test_choose_smoke_therapists_selects_multiple_low_load_real_providers():
    module = _load_tool_module()
    sessions = (
        _session("1", "A", 9),
        _session("2", "A", 10),
        _session("3", "A", 11),
        _session("4", "B", 9),
        _session("5", "C", 10),
        _session("6", "C", 11),
    )

    selected = module.choose_smoke_therapists(
        sessions,
        allowed_therapists=("A", "B", "C", "D"),
        therapist_count=2,
        min_sessions_per_therapist=1,
    )

    assert tuple(item[0] for item in selected) == ("B", "C")
    assert tuple(len(item[1]) for item in selected) == (1, 2)


def test_preparer_requires_true_multi_absence_and_full_app_contract():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert 'if therapist_count < 2:' in source
    assert 'assert_application_contract(source)' in source
    assert 'assert_application_contract(output)' in source
    assert '_write_absence_rows(preview, tuple(item[0] for item in selected))' in source
    assert 'ws.Cells(row, 2).Value = "ΝΑΙ"' in source
    assert 'Application contract output: OK' in source
