import inspect

from tools import prepare_daily_input_in_place as module


def test_daily_input_preparation_only_builds_validated_plan():
    source = inspect.getsource(module.prepare_daily_input_payload)

    assert "assert_application_contract(source)" in source
    assert "read_daily_input" in source
    assert "read_daily_replacement_facts" in source
    assert "build_daily_session_states" in source
    assert "build_patient_centric_preview_plan" in source
    assert "compose_outpatient_daily_plan" in source
    assert "validate_write_plan(plan)" in source
    assert "file_sha256(source)" in source

    assert "apply_write_plan_to_copy" not in source
    assert "shutil.copy" not in source
    assert "Workbooks.Open" not in source


def test_daily_input_effective_state_preserves_active_replacements():
    source = inspect.getsource(module.prepare_daily_input_payload)

    assert "replacements = read_daily_replacement_facts(" in source
    assert "replacements=replacements" in source
    assert '"active_replacements": len(replacements)' in source


def test_daily_input_payload_serializes_patch_semantics():
    source = inspect.getsource(module._patch_line)

    assert "patch.sheet" in source
    assert "patch.cell" in source
    assert "patch.intent.value" in source
    assert "patch.font_role" in source
    assert "patch.fill_role" in source
    assert "patch.border_role" in source
    assert "patch.wrap_text" in source
    assert "patch.min_font_size" in source
    assert "_text_runs_token(patch)" in source


def test_daily_input_payload_escape_contract():
    escaped = module._escape("A\\B\tC\nD\rE")
    assert escaped == "A\\\\B\\tC\\nD\\rE"
