import ast
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "Python" / "tools" / "apply_daily_input.py"


def _imported_names(source: str) -> dict[str, set[str]]:
    tree = ast.parse(source)
    result: dict[str, set[str]] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module:
            result.setdefault(node.module, set()).update(alias.name for alias in node.names)
    return result


def test_apply_daily_input_uses_enhanced_patient_registry():
    source = SCRIPT.read_text(encoding="utf-8")
    imports = _imported_names(source)
    assert "read_patient_registry" in imports.get("rehab_excel.patient_registry_source", set())
    assert "patients = read_patient_registry(input_book)" in source


def test_apply_daily_input_uses_unified_recurring_schedule():
    source = SCRIPT.read_text(encoding="utf-8")
    imports = _imported_names(source)
    assert "read_unified_base_schedule" in imports.get(
        "rehab_excel.outpatient_schedule_source", set()
    )
    assert "base_entries = read_unified_base_schedule(input_book)" in source
    assert "read_base_schedule(input_book)" not in source


def test_apply_daily_input_composes_outpatient_presentation_before_writeback():
    source = SCRIPT.read_text(encoding="utf-8")
    assert "compose_outpatient_daily_plan(" in source
    assert "report = apply_write_plan_to_copy(\n            composed_plan," in source
