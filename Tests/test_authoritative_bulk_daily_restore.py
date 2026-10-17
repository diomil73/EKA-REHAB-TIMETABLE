from __future__ import annotations

import inspect
from pathlib import Path
from types import SimpleNamespace

import pytest

from tools.restore_daily_replacement_facts import (
    AuthoritativeDailyRestoreError,
    _requested_restore_ids,
    restore_replacements_and_recompute,
)


def test_bulk_restore_pipeline_marks_set_before_single_recompute():
    source = inspect.getsource(restore_replacements_and_recompute)

    restore_pos = source.index("restore_daily_replacement_facts(stage, wanted_ids)")
    read_pos = source.index("read_daily_replacement_facts(")
    state_pos = source.index("build_daily_session_states(")

    assert restore_pos < read_pos < state_pos
    assert source.count("build_daily_session_states(") == 1
    assert source.count("write_daily_statistics(output, target_date, statistics)") == 1


def test_bulk_restore_pipeline_marks_each_ledger_row_only_after_recomputed_output_exists():
    source = inspect.getsource(restore_replacements_and_recompute)

    apply_pos = source.index("apply_write_plan_to_copy(")
    stats_pos = source.index("write_daily_statistics(output, target_date, statistics)")
    ledger_pos = source.index("for replacement_id in restored_ids:")

    assert apply_pos < stats_pos < ledger_pos
    assert "_mark_ledger_restored(output, replacement_id)" in source


def test_requested_restore_ids_uses_all_active_rows(monkeypatch):
    rows = (
        SimpleNamespace(replacement_id="R1"),
        SimpleNamespace(replacement_id="R2"),
    )
    monkeypatch.setattr(
        "tools.restore_daily_replacement_facts.read_daily_replacement_fact_rows",
        lambda *args, **kwargs: rows,
    )

    assert _requested_restore_ids(
        Path("book.xlsm"),
        target_date=None,
        explicit_ids=(),
        restore_all=True,
    ) == ("R1", "R2")


def test_requested_restore_ids_rejects_empty_explicit_set():
    with pytest.raises(AuthoritativeDailyRestoreError, match="At least one"):
        _requested_restore_ids(
            Path("book.xlsm"),
            target_date=None,
            explicit_ids=(),
            restore_all=False,
        )
