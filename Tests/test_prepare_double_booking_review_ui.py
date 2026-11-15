from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL = REPO_ROOT / "Python" / "tools" / "prepare_double_booking_review_ui.py"


def test_combined_ui_preparer_builds_review_and_rows_in_one_process():
    source = TOOL.read_text(encoding="utf-8")
    assert "build_double_booking_review_payload.main()" in source
    assert "export_review_rows(review, rows)" in source
    assert '"--review"' in source
    assert '"--rows"' in source
    assert "DOUBLE BOOKING REVIEW UI PREP OK" in source
    assert "subprocess" not in source
