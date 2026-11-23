"""Synthetic fixtures for strict, non-mutating full-workbook rollback audit."""
from shutil import copy2
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from rehab_excel.authoritative_commit import file_sha256
from rehab_excel.daily_input_full_rollback_audit import (
    FullRollbackAuditError,
    audit_full_workbook_rollback,
)
from rehab_excel.daily_input_journal import JOURNAL_HEADERS, JOURNAL_SHEET, _row
from rehab_excel.daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)


def fixture(tmp_path):
    source = tmp_path / "fictional-prepared.xlsm"
    restored = tmp_path / "fictional-restored.xlsm"
    operation = DailyInputOperation("FICTIONAL-ROLLBACK", "a" * 64, "2026-11-24")
    workbook = Workbook()
    daily = workbook.active
    daily.title = "DAILY_INPUT"
    daily["B2"] = "2026-11-24"
    master = workbook.create_sheet("MASTER_SCHEDULE")
    master["A1"] = "FICTIONAL CASE"
    master["B2"] = "=1+2"
    journal = workbook.create_sheet(JOURNAL_SHEET)
    journal.append(JOURNAL_HEADERS)
    journal.append(_row(operation))
    journal.sheet_state = "veryHidden"
    journal.protection.sheet = True
    workbook.save(source)
    workbook.close()
    copy2(source, restored)
    return source, restored, operation


def finish_rollback(restored, operation, *, with_rollback=True):
    failed = advance_daily_input_operation(operation, Phase.FAILED,
                                           failure_code="FICTIONAL_CRASH")
    workbook = load_workbook(restored, keep_vba=True)
    journal = workbook[JOURNAL_SHEET]
    journal.append(_row(failed))
    if with_rollback:
        rolled = advance_daily_input_operation(failed, Phase.ROLLED_BACK,
                                               rollback_verified=True)
        journal.append(_row(rolled))
    workbook.save(restored)
    workbook.close()


def audit(source, restored, *, digest=None, require_vba=False):
    return audit_full_workbook_rollback(
        source, restored,
        expected_snapshot_sha256=digest or file_sha256(source),
        require_vba=require_vba,
    )


def test_synthetic_rollback_preserves_all_nonjournal_package_members_exactly(tmp_path):
    source, restored, operation = fixture(tmp_path)
    before = source.read_bytes()
    finish_rollback(restored, operation)
    after = restored.read_bytes()
    receipt = audit(source, restored)
    assert receipt.checked_nonjournal_parts > 5
    assert receipt.original_journal_revisions == 1
    assert receipt.restored_journal_revisions == 3
    assert receipt.snapshot_sha256 == file_sha256(source)
    assert source.read_bytes() == before
    assert restored.read_bytes() == after


@pytest.mark.parametrize("mutation", ["value", "formula", "style", "resource"])
def test_business_cell_or_package_resource_change_fails_closed(tmp_path, mutation):
    source, restored, operation = fixture(tmp_path)
    finish_rollback(restored, operation)
    if mutation == "resource":
        with ZipFile(restored, "a") as archive:
            archive.writestr("xl/activeX/fakeControl.bin", b"DIFFERENT")
    else:
        workbook = load_workbook(restored, keep_vba=True)
        master = workbook["MASTER_SCHEDULE"]
        if mutation == "value":
            master["A1"] = "CHANGED FICTIONAL VALUE"
        elif mutation == "formula":
            master["B2"] = "=999"
        else:
            master["A1"].fill = PatternFill("solid", fgColor="FF0000")
        workbook.save(restored)
        workbook.close()
    with pytest.raises(FullRollbackAuditError, match="resource|inventory"):
        audit(source, restored)


def test_failed_without_verified_rollback_is_not_terminal_proof(tmp_path):
    source, restored, operation = fixture(tmp_path)
    finish_rollback(restored, operation, with_rollback=False)
    with pytest.raises(FullRollbackAuditError, match="prefix|revision|terminal"):
        audit(source, restored)


def test_snapshot_digest_must_match_exact_saved_bytes(tmp_path):
    source, restored, operation = fixture(tmp_path)
    finish_rollback(restored, operation)
    with pytest.raises(FullRollbackAuditError, match="snapshot changed"):
        audit(source, restored, digest="f" * 64)


def test_require_vba_refuses_fictional_nonmacro_fixture(tmp_path):
    source, restored, operation = fixture(tmp_path)
    finish_rollback(restored, operation)
    with pytest.raises(FullRollbackAuditError, match="ZIP|VBA"):
        audit(source, restored, require_vba=True)


def test_equal_paths_and_invalid_receipt_rejected_before_readback(tmp_path):
    source, restored, _ = fixture(tmp_path)
    with pytest.raises(FullRollbackAuditError, match="Distinct"):
        audit(source, source)
    with pytest.raises(FullRollbackAuditError, match="receipt"):
        audit_full_workbook_rollback(
            source, restored, expected_snapshot_sha256="invalid", require_vba=False,
        )


def test_unexplained_change_to_journal_history_rejected(tmp_path):
    source, restored, operation = fixture(tmp_path)
    finish_rollback(restored, operation)
    wb = load_workbook(restored, keep_vba=True)
    wb[JOURNAL_SHEET]["A2"] = "CHANGED-FICTIONAL-ID"
    wb.save(restored)
    wb.close()
    with pytest.raises(FullRollbackAuditError, match="journal|Journal"):
        audit(source, restored)
