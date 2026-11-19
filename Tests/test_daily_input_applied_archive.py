from dataclasses import replace
from datetime import date, time
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook

from rehab_core.models import Patient, Session
from rehab_excel.application_contract import DAILY_CHANGE_LEDGER_HEADERS
from rehab_excel.daily_input_applied_archive import (
    APPLIED_SHEET, HEADERS, applied_archive_rows, decode_applied_rows,
    plan_applied_archive_update, read_applied_archive, verify_applied_archive_readback,
)
from rehab_excel.daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError
from rehab_excel.daily_input_draft_archive import archive_rows, decode_archive_rows
from rehab_excel.daily_input_reader import interpret_daily_input_grid

DAY = date(2026, 10, 26)
NEXT = date(2026, 10, 27)


def fixture():
    today = DailyInputDayDraft(DAY, (("DEMO T", "ΝΑΙ", None, None, "Άδεια", None),), ((None,) * 5,))
    tomorrow = today.empty_for(NEXT)
    candidate = replace(today, patient_rows=(("DEMO P", "ΝΑΙ", None, "ΑΠΩΝ", "Εξέταση"),))
    patients = [Patient("P1", "DEMO P")]
    sessions = [Session("S1", "P1", "DEMO T", DAY, time(8, 30))]
    ledger = Workbook().active
    ledger.title = "DAILY_CHANGE_LEDGER"
    ledger.append(DAILY_CHANGE_LEDGER_HEADERS)
    for c in interpret_daily_input_grid(candidate, patients=patients, sessions=sessions).source_changes:
        ledger.append((c.change_id, DAY.isoformat(), c.change_type, c.subject_id,
                       c.display_name, "from", "to", c.reason, c.source_row, "ACTIVE", None, None))
    return dict(existing=(today, tomorrow), expected_previous=today, candidate=candidate,
                ledger_sheet=ledger, finalized_dates=(), patients=patients, sessions=sessions)


def sheet(rows):
    wb = Workbook()
    ws = wb.create_sheet(APPLIED_SHEET)
    ws.sheet_state = "veryHidden"
    ws.protection.sheet = True
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)
    return wb, ws


def test_apply_updates_only_its_date_and_exact_readback_survives_reopen():
    args = fixture()
    update = plan_applied_archive_update(**args)
    assert update.before == args["existing"]
    assert update.after == (args["candidate"], args["existing"][1])
    wb, ws = sheet(update.expected_rows)
    buffer = BytesIO()
    wb.save(buffer)
    wb.close()
    buffer.seek(0)
    reopened = load_workbook(buffer, data_only=False)
    try:
        verify_applied_archive_readback(reopened[APPLIED_SHEET], update)
    finally:
        reopened.close()


def test_applied_and_draft_storage_are_not_interchangeable():
    grids = fixture()["existing"]
    assert decode_applied_rows(applied_archive_rows(grids)) == grids
    with pytest.raises(DailyInputDaySwitchError):
        decode_applied_rows(archive_rows(grids))
    with pytest.raises(DailyInputDaySwitchError):
        decode_archive_rows(applied_archive_rows(grids))


@pytest.mark.parametrize("case", ["stale", "missing-prior", "final", "ledger", "duplicate", "layout"])
def test_invalid_apply_cannot_plan_archive_update(case):
    args = fixture()
    if case == "stale":
        args["expected_previous"] = args["candidate"]
    elif case == "missing-prior":
        args["expected_previous"] = None
    elif case == "final":
        args["finalized_dates"] = (DAY,)
    elif case == "ledger":
        args["ledger_sheet"].delete_rows(3)
    elif case == "duplicate":
        args["existing"] += (args["existing"][0],)
    else:
        candidate = args["candidate"]
        args["candidate"] = replace(candidate, patient_rows=candidate.patient_rows + ((None,) * 5,))
    with pytest.raises(DailyInputDaySwitchError):
        plan_applied_archive_update(**args)


@pytest.mark.parametrize("case", ["missing-date", "other-date-edit", "unprotected", "visible", "formula", "digest"])
def test_full_readback_rejects_missing_or_changed_records(case):
    update = plan_applied_archive_update(**fixture())
    rows = update.expected_rows
    if case == "missing-date":
        rows = applied_archive_rows(update.after[:1])
    elif case == "other-date-edit":
        other = replace(update.after[1], patient_rows=((None, None, None, None, "partial"),))
        rows = applied_archive_rows((update.after[0], other))
    wb, ws = sheet(rows)
    if case == "unprotected":
        ws.protection.sheet = False
    elif case == "visible":
        ws.sheet_state = "visible"
    elif case == "formula":
        ws["E2"] = "=1+1"
    elif case == "digest":
        ws["D2"] = "0" * 64
    try:
        with pytest.raises(DailyInputDaySwitchError):
            verify_applied_archive_readback(ws, update)
    finally:
        wb.close()
