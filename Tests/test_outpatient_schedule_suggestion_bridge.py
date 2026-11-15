from pathlib import Path

import pytest

from rehab_excel.outpatient_schedule_suggestion_bridge import (
    OutpatientScheduleSuggestionBridgeError,
    _parse_mode,
)
from rehab_excel.outpatient_schedule_suggestions import SuggestionMode


def test_bridge_accepts_mode_number_and_final_greek_label():
    assert _parse_mode(1) is SuggestionMode.CHANGE_TIME
    assert _parse_mode("6") is SuggestionMode.CHANGE_THERAPIST_AND_DAYS
    assert _parse_mode("Αλλαγή θεραπευτή, ώρας και ημερών") is SuggestionMode.CHANGE_THERAPIST_TIME_AND_DAYS


def test_bridge_rejects_invalid_mode():
    with pytest.raises(OutpatientScheduleSuggestionBridgeError, match="mode is invalid"):
        _parse_mode("κάτι άλλο")


def test_bridge_source_module_exposes_flattened_lines_for_vba():
    source = Path("Python/rehab_excel/outpatient_schedule_suggestion_bridge.py").read_text(encoding="utf-8")
    assert '"suggestion_lines": suggestion_lines' in source
    assert 'f"{row[\'therapist\']} | {row[\'time\']} | {row[\'days\']}"' in source
