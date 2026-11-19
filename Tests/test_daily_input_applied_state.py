from dataclasses import replace
from datetime import date, time

import pytest
from openpyxl import Workbook

from rehab_core.models import Patient, Session
from rehab_excel.application_contract import DAILY_CHANGE_LEDGER_HEADERS
from rehab_excel.daily_input_applied_state import verify_day_input
from rehab_excel.daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError
from rehab_excel.daily_input_reader import interpret_daily_input_grid

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
