from hashlib import sha256

from openpyxl import Workbook

from tools.audit_pilot_source_readiness import inspect_source, main


def fixture(tmp_path, *, prior_history=False, duplicate=False):
    path = tmp_path / "PROPOSED_PILOT_SOURCE.xlsx"
    wb = Workbook()
    wb.active.title = "MASTER_SCHEDULE"
    wb.create_sheet("DAILY_INPUT")
    ws = wb.create_sheet("PATIENTS")
    ws.append(("Id", "Room", "Name", "Infectious", "Status"))
    ws.append(("001", None, "Demo A", None, "ACTIVE"))
    ws.append(("002", None, "Demo B", None, "ACTIVE"))
    if duplicate:
        ws.append(("001", None, "Demo C", None, "ACTIVE"))
    wb.create_sheet("THERAPIST_DAILY")
    if prior_history:
        sh = wb.create_sheet("FINALIZED_DAY_SNAPSHOTS")
        sh.append(("SnapshotId", "Date"))
        sh.append(("TEST-SNAPSHOT", "2026-10-30"))
    wb.save(path)
    wb.close()
    return path


def test_intake_never_approves_macro_free_xlsx_and_preserves_hash(tmp_path):
    path = fixture(tmp_path)
    before = sha256(path.read_bytes()).hexdigest()
    report = inspect_source(path)
    assert report["operational_sheets_found"] == 4
    assert report["patient_registry_rows"] == 2
    assert report["history_rows"] == 0
    assert "NOT_MACRO_ENABLED_XLSM" in report["issues"]
    assert "NO_VBA_PROJECT" in report["issues"]
    assert sha256(path.read_bytes()).hexdigest() == before


def test_prior_test_history_is_flagged_not_silently_migrated(tmp_path):
    path = fixture(tmp_path, prior_history=True)
    report = inspect_source(path)
    assert report["history_rows"] == 1
    assert "EXISTING_TEST_OR_UNKNOWN_HISTORY_REQUIRES_ISOLATION" in report["issues"]
    assert report["patient_registry_rows"] == 2


def test_duplicate_patient_identifiers_block_pilot_candidate(tmp_path):
    path = fixture(tmp_path, duplicate=True)
    report = inspect_source(path)
    assert report["duplicate_patient_ids"] == 1
    assert "DUPLICATE_PATIENT_IDS" in report["issues"]


def test_missing_required_sheets_cannot_pass(tmp_path):
    path = fixture(tmp_path)
    from openpyxl import load_workbook
    book = load_workbook(path)
    del book["DAILY_INPUT"]
    book.save(path)
    book.close()
    report = inspect_source(path)
    assert report["operational_sheets_found"] == 3
    assert "DAILY_INPUT" in report["missing_operational_sheets"]
    assert "MISSING_OPERATIONAL_SHEETS" in report["issues"]


def test_missing_source_is_safe_stop(tmp_path):
    import sys
    from unittest.mock import patch
    with patch.object(sys, "argv", ["audit_pilot_source_readiness", "--workbook", str(tmp_path / "none.xlsm")]):
        assert main() == 2
