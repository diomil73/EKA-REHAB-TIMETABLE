from dataclasses import replace
from datetime import date, time
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

from rehab_core.models import Patient, Session
from rehab_excel.application_contract import DAILY_CHANGE_LEDGER_HEADERS
from rehab_excel.daily_input_applied_archive import APPLIED_SHEET, HEADERS as APPLIED_HEADERS
from rehab_excel.daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError, plan_open_day_switch
from rehab_excel.daily_input_draft_archive import ARCHIVE_SHEET, HEADERS as DRAFT_HEADERS, archive_rows
from rehab_excel.daily_input_reader import interpret_daily_input_grid
from rehab_excel.daily_input_verified_switch import verify_switch_inputs, verify_switch_storage_readback

DAY = date(2026, 10, 26)
NEXT = date(2026, 10, 27)


def scenario():
    today = DailyInputDayDraft(DAY, (("DEMO T1", "ΝΑΙ", None, None, "Άδεια", None),), ((None,) * 5,))
    tomorrow = DailyInputDayDraft(NEXT, ((None,) * 6,), (("DEMO P2", "ΝΑΙ", None, "ΑΠΩΝ", None),))
    today_draft = replace(today, patient_rows=(("DEMO P1", None, None, None, "partial"),))
    tomorrow_draft = replace(tomorrow, therapist_rows=(("DEMO T2", None, None, None, None, "partial"),))
    patients = [Patient("P1", "DEMO P1"), Patient("P2", "DEMO P2")]
    sessions = [Session("S1", "P1", "DEMO T1", DAY, time(8, 30)),
                Session("S2", "P2", "DEMO T2", NEXT, time(9))]
    ws = Workbook().active
    ws.title = "DAILY_CHANGE_LEDGER"
    ws.append(DAILY_CHANGE_LEDGER_HEADERS)
    for grid in (today, tomorrow):
        for c in interpret_daily_input_grid(grid, patients=patients, sessions=sessions).source_changes:
            ws.append((c.change_id, grid.service_date.isoformat(), c.change_type, c.subject_id,
                       c.display_name, "from", "to", c.reason, c.source_row, "ACTIVE", None, None))
    return today_draft, tomorrow_draft, (today, tomorrow), ws, patients, sessions


def arguments():
    current, target, applied, ledger, patients, sessions = scenario()
    plan = plan_open_day_switch(current=current, saved_drafts=(target,), expected_current_date=DAY,
                                target_date=NEXT, finalized_dates=())
    return dict(plan=plan, applied_grids=applied, ledger_sheet=ledger, patients=patients,
                sessions=sessions, therapist_names=(), replacement_record_dates=())


def workbook_for(expected):
    wb = Workbook()
    ws = wb.active
    ws.title = "DAILY_INPUT"
    ws["B2"] = expected.incoming.draft.service_date.isoformat()
    for header, headings, values in (
        (6, ("Θεραπευτής", "Όλη ημέρα", "Από", "Έως", "Αιτία", "Σχόλιο"), expected.incoming.draft.therapist_rows[0]),
        (11, ("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο"), expected.incoming.draft.patient_rows[0]),
    ):
        for col, value in enumerate(headings, 1):
            ws.cell(header, col, value)
        for col, value in enumerate(values, 1):
            ws.cell(header + 1, col, value)
    for name, area in (("PY_THERAPISTS", "A7"), ("PY_PATIENTS", "A12")):
        dv = DataValidation(type="list", formula1="=" + name)
        dv.add(area)
        ws.add_data_validation(dv)
    for name, headers, rows in ((ARCHIVE_SHEET, DRAFT_HEADERS, expected.draft_rows),
                                (APPLIED_SHEET, APPLIED_HEADERS, expected.applied_rows)):
        store = wb.create_sheet(name)
        store.sheet_state = "veryHidden"
        store.protection.sheet = True
        store.append(headers)
        for row in rows:
            store.append(row)
    return wb


def test_switch_preserves_both_drafts_but_projects_only_applied_inputs():
    args = arguments()
    verified = verify_switch_inputs(**args)
    assert verified.outgoing.has_unapplied_edits and verified.incoming.has_unapplied_edits
    assert verified.outgoing.applied_result.absences[0].subject_id == "DEMO T1"
    assert verified.incoming.applied_result.absences[0].subject_id == "P2"
    wb = workbook_for(verified)
    stream = BytesIO()
    wb.save(stream)
    wb.close()
    stream.seek(0)
    reopened = load_workbook(stream, data_only=False)
    try:
        verify_switch_storage_readback(reopened, verified)
    finally:
        reopened.close()


@pytest.mark.parametrize("case", ["outgoing-missing", "incoming-missing", "draft-missing", "ledger-mismatch"])
def test_missing_or_inconsistent_evidence_stops_switch(case):
    args = arguments()
    if case == "outgoing-missing":
        args["applied_grids"] = args["applied_grids"][1:]
    elif case == "incoming-missing":
        args["applied_grids"] = args["applied_grids"][:1]
    elif case == "draft-missing":
        args["plan"] = plan_open_day_switch(current=args["plan"].outgoing, saved_drafts=(),
            expected_current_date=DAY, target_date=NEXT, finalized_dates=())
    else:
        args["ledger_sheet"].delete_rows(3)
    with pytest.raises(DailyInputDaySwitchError):
        verify_switch_inputs(**args)


def test_fresh_day_initialization_is_explicit_and_checks_restored_history():
    args = arguments()
    new = date(2026, 10, 28)
    args["plan"] = plan_open_day_switch(current=args["plan"].outgoing, saved_drafts=(),
                                       expected_current_date=DAY, target_date=new, finalized_dates=())
    with pytest.raises(DailyInputDaySwitchError):
        verify_switch_inputs(**args)
    args["initialize_new_target"] = True
    actual = verify_switch_inputs(**args)
    assert actual.incoming.applied_grid == actual.incoming.draft
    assert not actual.incoming.applied_result.absences
    args["replacement_record_dates"] = (new,)
    with pytest.raises(DailyInputDaySwitchError):
        verify_switch_inputs(**args)
    args["replacement_record_dates"] = ()
    args["ledger_sheet"].append(("prior", new.isoformat(), "patient_absence", "P1", "DEMO P1",
                                 "from", "to", "reason", 12, "RESTORED", None, None))
    with pytest.raises(DailyInputDaySwitchError):
        verify_switch_inputs(**args)


@pytest.mark.parametrize("case", ["visible-date", "visible-grid", "draft-archive", "applied-archive", "protection"])
def test_post_save_mismatch_blocks_transaction(case):
    verified = verify_switch_inputs(**arguments())
    wb = workbook_for(verified)
    if case == "visible-date":
        wb["DAILY_INPUT"]["B2"] = DAY.isoformat()
    elif case == "visible-grid":
        wb["DAILY_INPUT"]["F7"] = "changed"
    elif case == "draft-archive":
        wb[ARCHIVE_SHEET].delete_rows(2)
    elif case == "applied-archive":
        wb[APPLIED_SHEET].delete_rows(2)
    else:
        wb[ARCHIVE_SHEET].protection.sheet = False
    try:
        with pytest.raises(DailyInputDaySwitchError):
            verify_switch_storage_readback(wb, verified)
    finally:
        wb.close()
