from __future__ import annotations

import inspect
from pathlib import Path
from types import SimpleNamespace

import pytest

from tools.restore_daily_replacement_facts import (
    AuthoritativeDailyRestoreError,
    _finalize_restore_output_once,
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


def test_bulk_restore_pipeline_finalizes_statistics_and_ledger_together():
    source = inspect.getsource(restore_replacements_and_recompute)

    apply_pos = source.index("apply_write_plan_to_copy(")
    finalize_pos = source.index("_finalize_restore_output_once(")

    assert apply_pos < finalize_pos
    assert "write_daily_statistics(output, target_date, statistics)" not in source
    assert "_mark_ledger_restored(output, replacement_id)" not in source
    assert "_mark_ledger_rows_restored_once(" not in source
    assert "replacement_ids=tuple(restored_ids)" in source


def test_bulk_output_finalization_uses_one_excel_session_for_stats_and_ledger():
    source = inspect.getsource(_finalize_restore_output_once)

    assert source.count('DispatchEx("Excel.Application")') == 1
    assert "build_daily_statistics_rows(target_date, statistics)" in source
    assert "for replacement_id in replacement_ids:" in source
    assert "STATUS_RESTORED" in source
    assert "RefreshDailyChangeLedgerView" in source
    assert "workbook.Save()" in source


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
