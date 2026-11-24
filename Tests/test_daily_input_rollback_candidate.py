"""Synthetic-only tests for non-destructive recovery candidate construction."""
from shutil import copy2
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.authoritative_commit import file_sha256
from rehab_excel.daily_input_full_rollback_audit import (
    audit_full_workbook_rollback,
    audit_restored_business_state_before_rollback,
    FullRollbackAuditError,
)
from rehab_excel.daily_input_journal import JOURNAL_HEADERS, JOURNAL_SHEET, _row
from rehab_excel.daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_rollback_candidate import (
    RollbackCandidateError,
    build_rollback_candidate,
    _journal_xml_for_history,
)


def _synthetic_files(tmp_path, *, append_failed=True, modify_business=True):
    snapshot = tmp_path / "FICTIONAL_PREPARED.xlsm"
    failed_book = tmp_path / "FICTIONAL_FAILED.xlsm"
    candidate = tmp_path / "FICTIONAL_RESTORED_CANDIDATE.xlsm"
    first = DailyInputOperation("FICTIONAL-OPERATION", "a" * 64, "2026-11-24")
    workbook = Workbook()
    workbook.active.title = "DAILY_INPUT"
    workbook.active["B2"] = "2026-11-24"
    master = workbook.create_sheet("MASTER_SCHEDULE")
    master["A1"] = "FICTIONAL ORIGINAL"
    master["B2"] = "=100+4"
    journal = workbook.create_sheet(JOURNAL_SHEET)
    journal.append(JOURNAL_HEADERS)
    journal.append(_row(first))
    journal.protection.sheet = True
    journal.sheet_state = "veryHidden"
    workbook.save(snapshot)
    workbook.close()
    copy2(snapshot, failed_book)
    if append_failed or modify_business:
        workbook = load_workbook(failed_book, keep_vba=True)
        if modify_business:
            workbook["MASTER_SCHEDULE"]["A1"] = "FICTIONAL UNFINISHED APPLY"
            workbook["MASTER_SCHEDULE"]["B2"] = "=0"
        if append_failed:
            failed = advance_daily_input_operation(
                first, Phase.FAILED, failure_code="FICTIONAL_APPLY_FAILURE",
            )
            workbook[JOURNAL_SHEET].append(_row(failed))
        workbook.save(failed_book)
        workbook.close()
    return snapshot, failed_book, candidate


def _build(snapshot, failed, candidate, *, digest=None, require_vba=False):
    return build_rollback_candidate(
        snapshot, failed, candidate,
        expected_snapshot_sha256=digest or file_sha256(snapshot),
        require_vba=require_vba,
    )


def test_new_candidate_restores_business_and_retains_failed_history(tmp_path):
    snapshot, failed, candidate = _synthetic_files(tmp_path)
    before_snap = snapshot.read_bytes()
    before_fail = failed.read_bytes()
    result = _build(snapshot, failed, candidate)
    assert candidate.exists()
    assert result.candidate_sha256 == file_sha256(candidate)
    assert result.nonjournal_parts_preserved > 5
    assert snapshot.read_bytes() == before_snap
    assert failed.read_bytes() == before_fail

    candidate_book = load_workbook(candidate, keep_vba=True)
    try:
        assert candidate_book["MASTER_SCHEDULE"]["A1"].value == "FICTIONAL ORIGINAL"
        assert candidate_book["MASTER_SCHEDULE"]["B2"].value == "=100+4"
        assert candidate_book[JOURNAL_SHEET].cell(3, 3).value == Phase.FAILED
        assert candidate_book[JOURNAL_SHEET].sheet_state == "veryHidden"
        assert candidate_book[JOURNAL_SHEET].protection.sheet
    finally:
        candidate_book.close()
    proof = audit_restored_business_state_before_rollback(
        snapshot, candidate, expected_snapshot_sha256=file_sha256(snapshot),
        require_vba=False,
    )
    assert proof.restored_journal_revisions == 2
    with pytest.raises(FullRollbackAuditError):
        audit_full_workbook_rollback(
            snapshot, candidate, expected_snapshot_sha256=file_sha256(snapshot),
            require_vba=False,
        )


def test_will_not_overwrite_existing_output(tmp_path):
    snapshot, failed, candidate = _synthetic_files(tmp_path)
    candidate.write_bytes(b"EXISTING USER FILE")
    with pytest.raises(RollbackCandidateError, match="unused XLSM"):
        _build(snapshot, failed, candidate)
    assert candidate.read_bytes() == b"EXISTING USER FILE"


def test_no_failed_journal_means_no_recovery_candidate(tmp_path):
    snapshot, failed, candidate = _synthetic_files(tmp_path, append_failed=False)
    with pytest.raises(RollbackCandidateError, match="FAILED"):
        _build(snapshot, failed, candidate)
    assert not candidate.exists()


def test_changed_prepared_snapshot_digest_refuses_without_output(tmp_path):
    snapshot, failed, candidate = _synthetic_files(tmp_path)
    with pytest.raises(RollbackCandidateError, match="Snapshot changed"):
        _build(snapshot, failed, candidate, digest="f" * 64)
    assert not candidate.exists()


def test_corrupt_failed_journal_prefix_is_rejected_without_output(tmp_path):
    snapshot, failed, candidate = _synthetic_files(tmp_path)
    wb = load_workbook(failed, keep_vba=True)
    wb[JOURNAL_SHEET]["A2"] = "ALTERED-FICTIONAL-ID"
    wb.save(failed)
    wb.close()
    with pytest.raises(RollbackCandidateError):
        _build(snapshot, failed, candidate)
    assert not candidate.exists()


def test_real_macro_project_is_mandatory_by_default(tmp_path):
    snapshot, failed, candidate = _synthetic_files(tmp_path)
    with pytest.raises(RollbackCandidateError):
        _build(snapshot, failed, candidate, require_vba=True)
    assert not candidate.exists()


def test_excel_namespaces_and_mc_ignorable_survive_journal_reconstruction():
    """Excel can reject an ET-rewritten sheet with dangling Ignorable prefixes."""
    original = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
        b'mc:Ignorable="x14ac"><dimension ref="A1:E2"/>'
        b'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c></row></sheetData>'
        b'<sheetProtection sheet="1"/><extLst/></worksheet>'
    )
    rows = (("FICTIVE & <operation>", 1, "PREPARED", "ok", "none"),)
    rebuilt = _journal_xml_for_history(original, rows)
    assert b'mc:Ignorable="x14ac"' in rebuilt
    assert b'xmlns:x14ac=' in rebuilt
    assert b'<sheetProtection sheet="1"/><extLst/>' in rebuilt
    assert rebuilt.startswith(original.split(b'<sheetData>')[0].replace(b'A1:E2', b'A1:E2'))
    assert b'FICTIVE &amp; &lt;operation&gt;' in rebuilt
    root = ET.fromstring(rebuilt)
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    assert root.find(f'{ns}sheetData/{ns}row/{ns}c/{ns}is/{ns}t').text == JOURNAL_HEADERS[0]
    assert root.find(f'{ns}dimension').get('ref') == 'A1:E2'


def test_prefixed_excel_sheetdata_is_preserved_without_rewriting_prefixes():
    original = (
        b'<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
        b'mc:Ignorable="xr"><x:dimension ref="A1:E1"/>'
        b'<x:sheetData><x:row r="1"/></x:sheetData><x:sheetProtection sheet="1"/>'
        b'</x:worksheet>'
    )
    updated = _journal_xml_for_history(original, ())
    assert b'<x:sheetData>' in updated
    assert b'<x:row r="1">' in updated
    assert b'xmlns:xr=' in updated
    assert b'mc:Ignorable="xr"' in updated
    assert b'<x:sheetProtection sheet="1"/>' in updated
    assert ET.fromstring(updated).tag.endswith('worksheet')


def test_journal_xml_rejects_ambiguous_sheetdata_markup():
    original = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<sheetData/><sheetProtection sheet="1"/></worksheet>'
    )
    with pytest.raises(RollbackCandidateError):
        _journal_xml_for_history(original, ())
