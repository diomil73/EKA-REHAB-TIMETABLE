import importlib.util
from datetime import date
from pathlib import Path

from rehab_excel.daily_replacement_facts import DailyReplacementFactRow
from rehab_core.models import ReplacementProviderKind, Session


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "apply_global_daily_replacements.py"


def _load_tool_module():
    spec = importlib.util.spec_from_file_location("apply_global_daily_replacements", TOOL_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_tool_uses_global_solver_and_not_sequential_choice_loop():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "build_global_replacement_assignments(" in source
    assert "persist_global_replacement_result(" in source
    assert "build_global_replacement_ledger_records(" in source
    assert "install_daily_change_ledger_preview(" in source
    assert "compose_master_daily_overlay(" in source
    assert "_ask_resolution(" not in source
    assert "Choose candidate" not in source
    assert "SEQUENTIAL THERAPIST ABSENCE RESOLUTION" not in source


def test_tool_resolves_global_solver_leftovers_as_persisted_treatment_suspensions():
    source = TOOL_PATH.read_text(encoding="utf-8")

    build_suspensions = source.index("generated_suspensions = build_unresolved_suspensions(")
    build_states = source.index("states = build_daily_session_states(")
    write_preview = source.index("report = apply_write_plan_to_copy")
    persist_replacements = source.index("persist_global_replacement_result(")
    persist_suspensions = source.index("persisted_suspensions = persist_daily_input_suspensions(")
    read_back = source.index("persisted_daily = read_daily_input(", persist_suspensions)

    assert build_suspensions < build_states < write_preview < persist_replacements < persist_suspensions < read_back
    assert "all_cancellations = daily.cancellations + generated_suspensions" in source
    assert "require_fully_resolved=False" in source[persist_replacements:persist_suspensions]
    assert "SessionCancellationKind.TREATMENT_SUSPENDED" in source
    assert 'print("Unresolved sessions: 0")' in source
    assert "Global replacement plan is incomplete; no preview will be written" not in source


def test_tool_requires_complete_app_contract_before_and_after_write():
    source = TOOL_PATH.read_text(encoding="utf-8")

    input_check = source.index("input_contract = assert_application_contract(input_book)")
    write_preview = source.index("report = apply_write_plan_to_copy")
    output_check = source.index("output_contract = assert_application_contract(output)")

    assert input_check < write_preview < output_check
    assert "ApplicationContractError" in source
    assert "Application contract input: OK" in source
    assert "Application contract output: OK" in source


def test_replacement_id_factory_allocates_stable_next_version():
    module = _load_tool_module()
    session = Session(
        session_id="S-1",
        patient_id="P-1",
        therapist_id="T-ABSENT",
        session_date=date(2026, 10, 21),
        start_time=__import__("datetime").time(9, 0),
    )
    history = (
        DailyReplacementFactRow(
            replacement_id="daily-global:S-1:2026-10-21",
            target_session_id="S-1",
            patient_id="P-1",
            original_therapist_id="T-ABSENT",
            replacement_therapist_id="T-1",
            replacement_date=date(2026, 10, 21),
            replacement_time=__import__("datetime").time(10, 0),
            reason="old",
            provider_kind=ReplacementProviderKind.THERAPIST,
            status="RESTORED",
        ),
        DailyReplacementFactRow(
            replacement_id="daily-global:S-1:2026-10-21:v2",
            target_session_id="S-1",
            patient_id="P-1",
            original_therapist_id="T-ABSENT",
            replacement_therapist_id="T-2",
            replacement_date=date(2026, 10, 21),
            replacement_time=__import__("datetime").time(11, 0),
            reason="old2",
            provider_kind=ReplacementProviderKind.THERAPIST,
            status="RESTORED",
        ),
    )

    factory = module._replacement_id_factory(history, {session.session_id: session})

    assert factory("S-1") == "daily-global:S-1:2026-10-21:v3"
    assert factory("S-1") == "daily-global:S-1:2026-10-21:v4"


def test_tool_loads_provider_policy_and_service_hours_before_global_solve():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert 'load_policy_book(REPO_ROOT / "Config" / "provider_policy.json")' in source
    assert 'load_service_hours(REPO_ROOT / "Config" / "service_hours.json")' in source
    assert "policy_book=policy_book" in source
    assert "service_hours=service_hours" in source


def test_tool_wires_reclining_adjacency_into_global_solve():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "build_reclining_adjacency_penalty" in source
    assert "adjacency_penalty = build_reclining_adjacency_penalty(base_entries)" in source
    assert "adjacency_penalty=adjacency_penalty" in source


def test_tool_wires_explicit_pool_service_ids_into_global_solve():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "build_service_id_by_session" in source
    assert "service_id_by_session = build_service_id_by_session(base_entries, target_date)" in source
    assert "service_id_by_session=service_id_by_session" in source


def test_tool_wires_authoritative_students_through_solve_and_projection():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "from rehab_excel.student_registry import StudentRegistryError, read_students" in source
    assert "students = tuple(read_students(input_book))" in source
    assert source.count("students=students") >= 3
    assert "StudentRegistryError" in source
