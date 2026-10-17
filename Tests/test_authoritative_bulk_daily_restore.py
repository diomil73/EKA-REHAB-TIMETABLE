from __future__ import annotations

import inspect
from pathlib import Path
from types import SimpleNamespace

import pytest

from tools.restore_daily_replacement_facts import (
    AuthoritativeDailyRestoreError,
    _RestoreOutputBackend,
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


def test_bulk_restore_pipeline_uses_single_output_backend_for_plan_stats_and_ledger():
    source = inspect.getsource(restore_replacements_and_recompute)

    apply_pos = source.index("apply_write_plan_to_copy(")
    backend_pos = source.index("backend=_RestoreOutputBackend(")

    assert apply_pos < backend_pos
    assert "write_daily_statistics(output, target_date, statistics)" not in source
    assert "_finalize_restore_output_once(" not in source
    assert "_mark_ledger_restored(output, replacement_id)" not in source
    assert "replacement_ids=tuple(restored_ids)" in source


def test_restore_output_backend_uses_one_excel_session_and_one_save():
    source = inspect.getsource(_RestoreOutputBackend.apply)

    assert source.count('DispatchEx("Excel.Application")') == 1
    assert "for patch in patches:" in source
    assert "self._write_statistics(workbook)" in source
    assert "self._mark_ledger_rows(workbook)" in source
    assert "RefreshDailyChangeLedgerView" in source
    assert source.count("workbook.Save()") == 1


def test_restore_output_backend_keeps_stats_and_ledger_helpers_in_same_backend():
    stats_source = inspect.getsource(_RestoreOutputBackend._write_statistics)
    ledger_source = inspect.getsource(_RestoreOutputBackend._mark_ledger_rows)

    assert "build_daily_statistics_rows(self.target_date, self.statistics)" in stats_source
    assert "for replacement_id in self.replacement_ids:" in ledger_source
    assert "STATUS_RESTORED" in ledger_source


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
