from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
PREPARE_SCRIPT = REPO_ROOT / "Python" / "tools" / "prepare_replacement_operator_decision.py"
OPERATOR_VBA = REPO_ROOT / "Python" / "rehab_excel" / "replacement_operator_vba.py"


def test_manual_replacement_preparation_does_not_run_global_solver() -> None:
    source = PREPARE_SCRIPT.read_text(encoding="utf-8")

    assert "build_global_replacement_assignments" not in source
    assert "build_reclining_adjacency_penalty" not in source
    assert '"is_global_recommendation": "0"' in source
    assert '"recommended_provider_id": ""' in source
    assert '"recommended_time": ""' in source


def test_automatic_replacement_route_still_uses_global_solver_path() -> None:
    source = OPERATOR_VBA.read_text(encoding="utf-8")

    assert "Case vbYes" in source
    assert "ResolveReplacementWorkflow = ResolveDailyReplacementsInPlace()" in source
    assert "Case vbNo" in source
    assert "ResolveReplacementWorkflow = RunManualReplacementWorkflow()" in source
