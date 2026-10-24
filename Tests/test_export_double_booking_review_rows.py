import importlib.util
import json
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "export_double_booking_review_rows.py"


def _load_tool_module():
    spec = importlib.util.spec_from_file_location("export_double_booking_review_rows", TOOL_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_export_review_rows_flattens_ui_choices_without_exposing_patient_ids(tmp_path):
    module = _load_tool_module()
    review = tmp_path / "review.json"
    output = tmp_path / "rows.tsv"
    review.write_text(
        json.dumps(
            {
                "schema": "eka.double_booking_review.v1",
                "requires_user_confirmation": True,
                "items": [
                    {
                        "target_session_id": "S-1",
                        "patient_id": "P-SECRET",
                        "patient_name": "ΑΣΘΕΝΗΣ",
                        "original_time": "08:30",
                        "options": [
                            {
                                "therapist_id": "T-1",
                                "therapist_name": "ΘΕΡΑΠΕΥΤΗΣ",
                                "replacement_time": "08:30",
                                "warning": "ΠΡΟΕΙΔΟΠΟΙΗΣΗ",
                            }
                        ],
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    assert module.export_review_rows(review, output) == 1
    text = output.read_text(encoding="utf-8-sig")
    assert "ΑΣΘΕΝΗΣ" in text
    assert "ΘΕΡΑΠΕΥΤΗΣ" in text
    assert "P-SECRET" not in text
    assert "S-1" in text
    assert "T-1" in text


def test_export_review_rows_rejects_payload_without_explicit_confirmation(tmp_path):
    module = _load_tool_module()
    review = tmp_path / "review.json"
    output = tmp_path / "rows.tsv"
    review.write_text(
        json.dumps(
            {
                "schema": "eka.double_booking_review.v1",
                "requires_user_confirmation": False,
                "items": [],
            }
        ),
        encoding="utf-8",
    )

    try:
        module.export_review_rows(review, output)
    except ValueError as exc:
        assert "explicit confirmation" in str(exc)
    else:
        raise AssertionError("expected ValueError")


def test_ui_rows_bridge_is_read_only_with_respect_to_workbooks():
    source = TOOL_PATH.read_text(encoding="utf-8")
    assert "openpyxl" not in source
    assert "win32com" not in source
    assert "append_daily_replacement_facts" not in source
    assert "apply_write_plan_to_copy" not in source
