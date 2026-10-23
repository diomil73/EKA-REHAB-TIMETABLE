from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "build_double_booking_review_payload.py"


def test_review_payload_bridge_is_review_only_and_requires_confirmation():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "build_global_replacement_assignments(" in source
    assert "build_double_booking_review_items(" in source
    assert '"requires_user_confirmation": True' in source
    assert "Normal global plan has no unresolved sessions" in source
    assert "DOUBLE BOOKING REVIEW PAYLOAD OK" in source
    assert "Explicit user confirmation required: True" in source

    assert "apply_write_plan_to_copy(" not in source
    assert "persist_global_replacement_result(" not in source
    assert "write_daily_statistics(" not in source
    assert "install_daily_change_ledger_preview(" not in source
    assert "confirm_double_booking_exception(" not in source


def test_review_payload_schema_contains_ui_fields():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert '"schema": "eka.double_booking_review.v1"' in source
    assert '"patient_name": item.patient_name' in source
    assert '"original_therapist_id": item.original_therapist_id' in source
    assert '"original_time": item.original_time.strftime("%H:%M")' in source
    assert '"therapist_name": option.therapist_name' in source
    assert '"replacement_time": option.replacement_time.strftime("%H:%M")' in source
    assert '"warning": option.warning' in source
