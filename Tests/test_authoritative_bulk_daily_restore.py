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


def test_bulk_restore_validates_set_before_single_recompute_without_stage_excel_mutation():
    source = inspect.getsource(restore_replacements_and_recompute)

    facts_pos = source.index("fact_rows = read_daily_replacement_fact_rows(")
    validate_pos = source.index("validate_batch_restore_request(fact_rows, wanted_ids)")
    derive_pos = source.index("_active_replacements_after_restore(")
    state_pos = source.index("build_daily_session_states(")

    assert facts_pos < validate_pos < derive_pos < state_pos
    assert "restore_daily_replacement_facts(stage, wanted_ids)" not in source
    assert "read_daily_replacement_facts(" not in source
    assert source.count("build_daily_session_states(") == 1


def test_bulk_restore_requires_application_contract_before_and_after_write():
    source = inspect.getsource(restore_replacements_and_recompute)

    input_contract_pos = source.index("assert_application_contract(source)")
    apply_pos = source.index("apply_write_plan_to_copy(")
    output_contract_pos = source.index("assert_application_contract(output)")

    assert input_contract_pos < apply_pos < output_contract_pos
    assert source.count("assert_application_contract(") == 2


def test_bulk_restore_pipeline_uses_single_output_backend_for_master_plan_stats_facts_and_ledger():
    source = inspect.getsource(restore_replacements_and_recompute)

    projection_pos = source.index("master_projection = build_master_projection(stage, patients=patients)")
    apply_pos = source.index("apply_write_plan_to_copy(")
    backend_pos = source.index("backend=_RestoreOutputBackend(")

    assert projection_pos < apply_pos < backend_pos
    assert "refresh_master_projection_in_place(stage)" not in source
    assert "write_daily_statistics(output, target_date, statistics)" not in source
    assert "_finalize_restore_output_once(" not in source
    assert "_mark_ledger_restored(output, replacement_id)" not in source
    assert "replacement_ids=tuple(restored_ids)" in source
    assert "master_projection=master_projection" in source


def test_bulk_restore_reuses_loaded_patients_and_master_projection():
    source = inspect.getsource(restore_replacements_and_recompute)

    assert "base_entries = read_unified_base_schedule(stage, patients=patients)" in source
    assert "master_projection_rows=master_projection" in source
    assert source.count("master_projection = build_master_projection(stage, patients=patients)") == 1


def test_restore_output_backend_uses_one_excel_session_and_one_save():
    source = inspect.getsource(_RestoreOutputBackend.apply)

    assert source.count('DispatchEx("Excel.Application")') == 1
    assert "self._apply_master_projection(workbook)" in source
    assert "for patch in patches:" in source
    assert "self._write_statistics(workbook)" in source
    assert "self._mark_replacement_fact_rows(workbook)" in source
    assert "self._mark_ledger_rows(workbook)" in source
    assert "RefreshDailyChangeLedgerView" in source
    assert source.count("workbook.Save()") == 1


def test_restore_output_backend_marks_authoritative_fact_rows_in_same_session():
    source = inspect.getsource(_RestoreOutputBackend._mark_replacement_fact_rows)

    assert "REPLACEMENT_FACTS_SHEET" in source
    assert "REPLACEMENT_FACT_HEADERS" in source
    assert "facts.Cells(row, 10).Value = STATUS_RESTORED" in source
    assert "facts.Visible = 2" in source
    assert "Missing replacement fact id(s)" in source


def test_restore_output_backend_rebuilds_master_with_batched_com_writes():
    source = inspect.getsource(_RestoreOutputBackend._apply_master_projection)

    assert "Win32ComMasterProjectionBackend(palette=self.palette)" in source
    assert 'workbook.Worksheets("MASTER_SCHEDULE")' in source
    assert "backend._template_formulas(ws)" in source
    assert 'ws.Range(f"A{template_row}:P{clear_last}").ClearContents()' in source
    assert ").Formula = tuple(formula_matrix)" in source
    assert "item_by_row = {item.target_row: item" in source
    assert "ws.Cells(row, col).Formula" not in source
    assert 'ws.Range(f"A{row}:P{row}").ClearContents()' not in source
    assert "if item.responsible_doctor:" in source
    assert "backend._render_patient_identity" in source
    assert "backend._apply_infectious_style" in source


def test_restore_output_backend_keeps_stats_and_ledger_helpers_in_same_backend():
    stats_source = inspect.getsource(_RestoreOutputBackend._write_statistics)
    ledger_source = inspect.getsource(_RestoreOutputBackend._mark_ledger_rows)

    assert "build_daily_statistics_rows(self.target_date, self.statistics)" in stats_source
    assert "for replacement_id in self.replacement_ids:" in ledger_source
    assert "STATUS_RESTORED" in ledger_source


def test_requested_restore_ids_uses_all_active_rows(monkeypatch):
    rows = (
        SimpleNamespace(replacement_id="R1", is_active=True),
        SimpleNamespace(replacement_id="R2", is_active=True),
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
