import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TOOLS = (
    ROOT / "Python" / "tools" / "apply_daily_input_replacements.py",
    ROOT / "Python" / "tools" / "apply_replacement_choice.py",
)


def _imports(path: Path) -> set[tuple[str, str]]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    found: set[tuple[str, str]] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module:
            for alias in node.names:
                found.add((node.module, alias.name))
    return found


def _source(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_replacement_tools_use_enhanced_patient_registry():
    for tool in TOOLS:
        imports = _imports(tool)
        assert (
            "rehab_excel.patient_registry_source",
            "read_patient_registry",
        ) in imports, tool
        assert "read_patient_registry(" in _source(tool)


def test_replacement_tools_use_unified_recurring_source():
    for tool in TOOLS:
        imports = _imports(tool)
        assert (
            "rehab_excel.outpatient_schedule_source",
            "read_unified_base_schedule",
        ) in imports, tool
        source = _source(tool)
        assert "read_unified_base_schedule(" in source
        assert "read_base_schedule(" not in source


def test_replacement_tools_compose_outpatient_presentation():
    for tool in TOOLS:
        imports = _imports(tool)
        assert (
            "rehab_excel.daily_preview_composer",
            "compose_outpatient_daily_plan",
        ) in imports, tool
        source = _source(tool)
        assert "compose_outpatient_daily_plan(" in source
        assert "apply_write_plan_to_copy(\n            composed_plan," in source
