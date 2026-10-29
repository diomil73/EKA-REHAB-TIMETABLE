from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
PREPARER = REPO_ROOT / "Python" / "tools" / "prepare_suspension_resolution_smoke.py"
VERIFIER = REPO_ROOT / "Python" / "tools" / "verify_suspension_resolution_smoke.py"


def test_suspension_smoke_preparer_forces_real_robotic_absence_scenario():
    source = PREPARER.read_text(encoding="utf-8")

    assert "choose_robotic_smoke_target" in source
    assert "session.robotic" in source
    assert 'daily_ws.Cells(7, 1).Value = absent_therapist_id' in source
    assert 'daily_ws.Cells(7, 2).Value = "ΝΑΙ"' in source
    assert 'students_ws.Cells(row, 7).Value = "ΟΧΙ"' in source
    assert "create_daily_input_preview" in source
    assert "Source unchanged" in source
    assert "assert_application_contract" in source


def test_suspension_smoke_verifier_requires_persisted_fact_and_finalized_state():
    source = VERIFIER.read_text(encoding="utf-8")

    assert "SessionCancellationKind.TREATMENT_SUSPENDED" in source
    assert "read_daily_input(" in source
    assert "read_daily_replacement_facts(" in source
    assert "build_daily_session_states(" in source
    assert "calculate_daily_treatment_statistics(" in source
    assert "statistics.unresolved_sessions != 0" in source
    assert "statistics.treatment_suspended_sessions < 1" in source
    assert "SUSPENSION RESOLUTION SMOKE VERIFIED" in source
