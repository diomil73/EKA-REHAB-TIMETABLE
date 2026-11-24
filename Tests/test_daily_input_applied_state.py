from dataclasses import replace
from datetime import date, time

import pytest
from openpyxl import Workbook, load_workbook

from rehab_core.models import Patient, Session
from rehab_excel.application_contract import DAILY_CHANGE_LEDGER_HEADERS
from rehab_excel.daily_input_applied_state import verify_day_input, verify_input_ledger_saved_readback
from rehab_excel.daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError
from rehab_excel.daily_input_reader import interpret_daily_input_grid, read_daily_input

DAY = date(2026, 10, 26)
NEXT = date(2026, 10, 27)


def fixture():
    applied = DailyInputDayDraft(
        DAY, (("DEMO T", "ΟΧΙ", "08:00", "10:00", "Άδεια", None),),
        (("DEMO P", "ΝΑΙ", None, "ΑΠΩΝ", None),),
    )
    patients = [Patient("P1", "DEMO P")]
    sessions = [Session("S1", "P1", "DEMO T", DAY, time(8, 30))]
    args = dict(patients=patients, sessions=sessions)
    result = interpret_daily_input_grid(applied, **args)
    ws = Workbook().active
    ws.title = "DAILY_CHANGE_LEDGER"
    ws.append(DAILY_CHANGE_LEDGER_HEADERS)
    for c in result.source_changes:
        ws.append((c.change_id, DAY.isoformat(), c.change_type, c.subject_id,
                   c.display_name, "from", "to", c.reason, c.source_row, "ACTIVE", None, None))
    ws.append(("other-day", NEXT.isoformat(), "patient_absence", "P2",
               "DEMO P2", "from", "to", "reason", 12, "ACTIVE", None, None))
    return applied, ws, args


def test_pending_incomplete_draft_does_not_change_applied_facts():
    applied, ws, args = fixture()
    draft = replace(applied, patient_rows=(("DEMO P", None, None, None, "partial"),))
    before = tuple(ws.values)
    verified = verify_day_input(draft=draft, applied_grid=applied, ledger_sheet=ws, **args)
    assert verified.has_unapplied_edits
    assert verified.draft == draft
    assert verified.applied_result == interpret_daily_input_grid(applied, **args)
    assert verified.applied_result.absences[0].end_time == time(10)
    assert tuple(ws.values) == before


def test_clean_flag_uses_exact_typed_grid_and_ignores_other_dates():
    applied, ws, args = fixture()
    assert not verify_day_input(draft=applied, applied_grid=applied, ledger_sheet=ws, **args).has_unapplied_edits
    empty = applied.empty_for(DAY)
    for row in (2, 3):
        ws.cell(row, 10).value = "RESTORED"
    clean = verify_day_input(draft=applied, applied_grid=empty, ledger_sheet=ws, **args)
    assert clean.has_unapplied_edits
    assert not clean.applied_result.source_changes


@pytest.mark.parametrize("case", ["missing", "wrong-date", "wrong-layout", "missing-ledger",
                                  "duplicate", "subject", "reason", "row", "status", "schema"])
def test_unproven_applied_state_fails_closed(case):
    applied, ws, args = fixture()
    stored = applied
    if case == "missing":
        stored = None
    elif case == "wrong-date":
        stored = replace(applied, service_date=NEXT)
    elif case == "wrong-layout":
        stored = replace(applied, patient_rows=applied.patient_rows + ((None,) * 5,))
    elif case == "missing-ledger":
        ws.delete_rows(2)
    elif case == "duplicate":
        ws.append([c.value for c in ws[2]])
    elif case == "subject":
        ws["D2"] = "DIFFERENT"
    elif case == "reason":
        ws["H2"] = "different reason"
    elif case == "row":
        ws["I2"] = 15
    elif case == "status":
        ws["J2"] = "UNKNOWN"
    else:
        ws["A1"] = "WRONG"
    before = tuple(ws.values)
    with pytest.raises(DailyInputDaySwitchError):
        verify_day_input(draft=applied, applied_grid=stored, ledger_sheet=ws, **args)
    assert tuple(ws.values) == before

def _saved_input_and_ledger_fixture(tmp_path):
    """Fictional native-shaped OOXML data and a matching real ledger worksheet."""
    source = tmp_path / "FICTIONAL_INPUT_LEDGER.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "DAILY_INPUT"
    ws["A2"] = "Ημερομηνία"
    ws["B2"] = "26/10/2026"
    for col, value in enumerate(
        ("Θεραπευτής", "Όλη ημέρα", "Από", "Έως", "Αιτία", "Σχόλιο"), 1
    ):
        ws.cell(6, col).value = value
    for col, value in enumerate(
        ("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο"), 1
    ):
        ws.cell(20, col).value = value
    ws["A21"] = "ΔΟΚΙΜΑΣΤΙΚΟΣ ΑΣΘΕΝΗΣ"
    ws["B21"] = "ΟΧΙ"
    ws["C21"] = "08:30"
    ws["D21"] = "ΑΠΩΝ"
    ws["E21"] = "ΣΥΝΘΕΤΙΚΗ ΑΙΤΙΑ"
    ledger = wb.create_sheet("DAILY_CHANGE_LEDGER")
    ledger.append(DAILY_CHANGE_LEDGER_HEADERS)
    ledger.sheet_state = "veryHidden"
    wb.save(source)
    wb.close()

    patients = (Patient("FICTIONAL-P", "ΔΟΚΙΜΑΣΤΙΚΟΣ ΑΣΘΕΝΗΣ"),)
    sessions = (Session("FICTIONAL-S", "FICTIONAL-P", "FICTIONAL-T", DAY, time(8, 30)),)
    result = read_daily_input(source, patients=patients, sessions=sessions)
    assert len(result.source_changes) == 1
    wb = load_workbook(source, keep_vba=True)
    ledger = wb["DAILY_CHANGE_LEDGER"]
    for change in result.source_changes:
        ledger.append((
            change.change_id, DAY.isoformat(), change.change_type,
            change.subject_id, change.display_name, "Πριν", "Μετά",
            change.reason, change.source_row, "ACTIVE", None, None,
        ))
    wb.save(source)
    wb.close()
    return source, patients, sessions, result.source_changes


def test_saved_input_ledger_reconciles_exact_preflight_membership_and_is_readonly(tmp_path):
    from hashlib import sha256

    source, patients, sessions, changes = _saved_input_and_ledger_fixture(tmp_path)
    before = source.read_bytes()
    result = verify_input_ledger_saved_readback(
        source, service_date=DAY, validated_changes=changes,
        patients=patients, sessions=sessions,
    )
    assert result.service_date == DAY
    assert result.source_change_count == 1
    assert result.active_ledger_count == 1
    assert result.workbook_sha256 == sha256(before).hexdigest()
    assert source.read_bytes() == before


@pytest.mark.parametrize("mutation", [
    "wrong-subject", "wrong-reason", "wrong-source-row", "unknown-status",
    "missing-row", "duplicate-active", "visible-ledger", "missing-ledger",
    "stale-input", "wrong-input-date",
])
def test_saved_input_ledger_rejects_corruption_and_keeps_source_bytes(
    tmp_path, mutation,
):
    source, patients, sessions, changes = _saved_input_and_ledger_fixture(tmp_path)
    wb = load_workbook(source, keep_vba=True)
    ledger = wb["DAILY_CHANGE_LEDGER"]
    if mutation == "wrong-subject":
        ledger["D2"] = "WRONG-PATIENT"
    elif mutation == "wrong-reason":
        ledger["H2"] = "UNAPPROVED"
    elif mutation == "wrong-source-row":
        ledger["I2"] = 999
    elif mutation == "unknown-status":
        ledger["J2"] = "UNKNOWN"
    elif mutation == "missing-row":
        ledger.delete_rows(2)
    elif mutation == "duplicate-active":
        ledger.append([cell.value for cell in ledger[2]])
    elif mutation == "visible-ledger":
        ledger.sheet_state = "visible"
    elif mutation == "missing-ledger":
        del wb["DAILY_CHANGE_LEDGER"]
    elif mutation == "stale-input":
        wb["DAILY_INPUT"]["E21"] = "ΑΛΛΑΞΕ ΜΕΤΑ ΤΟ PREFLIGHT"
    elif mutation == "wrong-input-date":
        wb["DAILY_INPUT"]["B2"] = "27/10/2026"
    wb.save(source)
    wb.close()
    before = source.read_bytes()
    with pytest.raises(DailyInputDaySwitchError):
        verify_input_ledger_saved_readback(
            source, service_date=DAY, validated_changes=changes,
            patients=patients, sessions=sessions,
        )
    assert source.read_bytes() == before


def test_saved_input_ledger_requires_prior_preflight_not_empty_guess(tmp_path):
    source, patients, sessions, changes = _saved_input_and_ledger_fixture(tmp_path)
    for invalid in ((), list(changes)):
        with pytest.raises(DailyInputDaySwitchError, match="preflight"):
            verify_input_ledger_saved_readback(
                source, service_date=DAY, validated_changes=invalid,
                patients=patients, sessions=sessions,
            )


def test_saved_input_ledger_detects_concurrent_workbook_change(tmp_path, monkeypatch):
    from rehab_excel import daily_input_applied_state as module
    from rehab_excel.authoritative_commit import file_sha256 as real_sha

    source, patients, sessions, changes = _saved_input_and_ledger_fixture(tmp_path)
    original = source.read_bytes()
    reads = 0

    def changing_digest(path):
        nonlocal reads
        reads += 1
        return real_sha(path) if reads == 1 else "0" * 64

    monkeypatch.setattr(module, "file_sha256", changing_digest)
    with pytest.raises(DailyInputDaySwitchError, match="changed"):
        verify_input_ledger_saved_readback(
            source, service_date=DAY, validated_changes=changes,
            patients=patients, sessions=sessions,
        )
    assert reads == 2
    assert source.read_bytes() == original


def test_saved_input_ledger_is_not_automatic_native_apply_authorization():
    from inspect import getsource
    import rehab_excel.daily_input_applied_state as module

    source = getsource(module.verify_input_ledger_saved_readback)
    assert "read_daily_input" in source
    assert "active_input_ledger_keys" in source
    assert "file_sha256" in source
    assert "ApplyDailyInputPreview" not in source
    assert "DailyInputJournalApplyEnabled = True" not in source
