from __future__ import annotations

from types import SimpleNamespace

import pytest

from rehab_excel.daily_replacement_fact_batch import validate_batch_restore_request
from rehab_excel.daily_replacement_facts import DailyReplacementFactError


def _fact(replacement_id: str, status: str = "ACTIVE"):
    return SimpleNamespace(replacement_id=replacement_id, status=status)


def test_batch_restore_accepts_only_when_entire_requested_set_is_active():
    facts = (_fact("R1"), _fact("R2"), _fact("R3", "RESTORED"))

    assert validate_batch_restore_request(facts, ("R1", "R2")) == ("R1", "R2")


def test_batch_restore_rejects_missing_member_before_mutation():
    facts = (_fact("R1"), _fact("R2"))

    with pytest.raises(DailyReplacementFactError, match="R3.*not found"):
        validate_batch_restore_request(facts, ("R1", "R3"))


def test_batch_restore_rejects_already_restored_member_before_mutation():
    facts = (_fact("R1"), _fact("R2", "RESTORED"))

    with pytest.raises(DailyReplacementFactError, match="R2.*already restored"):
        validate_batch_restore_request(facts, ("R1", "R2"))


def test_batch_restore_rejects_duplicate_request_ids():
    facts = (_fact("R1"),)

    with pytest.raises(DailyReplacementFactError, match="Duplicate ReplacementId"):
        validate_batch_restore_request(facts, ("R1", "R1"))
