import importlib.util
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "Python" / "tools" / "preview_operational_replacement.py"
spec = importlib.util.spec_from_file_location("preview_operational_replacement", SCRIPT)
module = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(module)


def test_append_unique_line_keeps_existing_once():
    assert module.append_unique_line("Α\nΒ", "Β") == "Α\nΒ"


def test_append_unique_line_adds_new_line():
    assert module.append_unique_line("Α", "Β") == "Α\nΒ"


def test_line_run_targets_full_line_containing_patient():
    text = "ΑΛΛΟΣ\nΖΑΛΟΚΩΣΤΑΣ [Καθ/να]"
    run = module.line_run_for_text(
        text,
        "ΖΑΛΟΚΩΣΤΑΣ",
        strike=True,
        font_role="muted",
    )
    assert run.start == len("ΑΛΛΟΣ\n") + 1
    assert run.length == len("ΖΑΛΟΚΩΣΤΑΣ [Καθ/να]")
    assert run.strike_through is True
