from pathlib import Path
import sys

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_excel.daily_input_reader import DailyInputReadError, _provider_id


def test_provider_id_accepts_official_registry_therapist_without_session_that_day():
    assert _provider_id((), "Πέτσιος", ("Πέτσιος", "Μηλιδάκης")) == "Πέτσιος"


def test_provider_id_still_rejects_unknown_therapist():
    with pytest.raises(DailyInputReadError, match="official registry/date"):
        _provider_id((), "Αγνωστος", ("Πέτσιος", "Μηλιδάκης"))
