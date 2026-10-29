import importlib.util
from datetime import date, time
from pathlib import Path

from rehab_core.models import Session


REPO_ROOT = Path(__file__).resolve().parents[1]
PREPARE_PATH = REPO_ROOT / "Python" / "tools" / "prepare_student_replacement_smoke.py"
VERIFY_PATH = REPO_ROOT / "Python" / "tools" / "verify_student_replacement_smoke.py"


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _session(
    session_id: str,
    therapist_id: str,
    hour: int,
    *,
    robotic: bool = False,
):
    return Session(
        session_id=session_id,
        patient_id=f"P-{session_id}",
        therapist_id=therapist_id,
        session_date=date(2026, 10, 26),
        start_time=time(hour, 0),
        robotic=robotic,
    )


def test_choose_student_smoke_target_prefers_small_non_robotic_load():
    module = _load(PREPARE_PATH, "prepare_student_replacement_smoke")
    sessions = (
        _session("A1", "A", 9),
        _session("A2", "A", 10),
        _session("B1", "B", 9),
        _session("C1", "C", 11, robotic=True),
    )

    therapist_id, affected = module.choose_student_smoke_target(
        sessions,
        allowed_therapists=("A", "B", "C"),
        max_displaced=5,
    )

    assert therapist_id == "B"
    assert tuple(item.session_id for item in affected) == ("B1",)


def test_prepare_tool_builds_full_app_student_smoke_contract():
    source = PREPARE_PATH.read_text(encoding="utf-8")

    assert 'SMOKE_STUDENT_ID = "0000-STUDENT-SMOKE"' in source
    assert "create_daily_input_preview(" in source
    assert "Win32ComStudentRegistrySchemaBackend().ensure_schema(preview)" in source
    assert "read_students(preview)" in source
    assert "read_daily_input(" in source
    assert "assert_application_contract(source)" in source
    assert "assert_application_contract(preview)" in source
    assert 'daily_ws.Cells(absence_row, 2).Value = "ΝΑΙ"' in source
    assert "Source unchanged:" in source


def test_verifier_requires_student_fact_and_therapist_daily_projection():
    source = VERIFY_PATH.read_text(encoding="utf-8")

    assert "ReplacementProviderKind.STUDENT" in source
    assert 'str(row.status).strip().upper() == "ACTIVE"' in source
    assert '_therapist_daily_contains(workbook_path, student.display_name)' in source
    assert 'print("STUDENT REPLACEMENT SMOKE VERIFIED")' in source
    assert 'print("THERAPIST_DAILY projection: OK")' in source
