from datetime import date, time

import pytest
from openpyxl import Workbook

from rehab_core.models import AbsenceKind, Patient, Session
from rehab_excel.daily_input_day_switch import DailyInputDayDraft
from rehab_excel.daily_input_reader import (
    DailyInputReadError, interpret_daily_input_grid, read_daily_input,
)

DAY = date(2026, 10, 26)
NEXT = date(2026, 10, 27)


def fixture(day=DAY):
    patients = [Patient("P1", "DEMO P1"), Patient("P2", "DEMO P2")]
    sessions = [Session("S1", "P1", "DEMO T", day, time(8, 30)),
                Session("S2", "P2", "DEMO T", day, time(9))]
    grid = DailyInputDayDraft(
        day,
        (("DEMO T", "ΟΧΙ", "08:00", "10:00", "Άδεια", "comment"),
         (None,) * 6),
        (("ΕΞ | DEMO P1", "ΝΑΙ", None, "ΑΠΩΝ", "comment"),
         ("DEMO P2", "ΟΧΙ", "09:00", "ΑΝΑΒΟΛΗ ΤΜΗΜΑΤΟΣ", None)),
    )
    return grid, patients, sessions


def test_literal_grid_uses_same_row_ids_and_facts_as_worksheet(tmp_path):
    grid, patients, sessions = fixture()
    wb = Workbook()
    ws = wb.active
    ws.title = "DAILY_INPUT"
    ws["B2"] = grid.service_date.strftime("%d/%m/%Y")
    for header, headings, rows in (
        (6, ("Θεραπευτής", "Όλη ημέρα", "Από", "Έως", "Αιτία", "Σχόλιο"), grid.therapist_rows),
        (12, ("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο"), grid.patient_rows),
    ):
        for col, value in enumerate(headings, 1):
            ws.cell(header, col, value)
        for row, values in enumerate(rows, header + 1):
            for col, value in enumerate(values, 1):
                ws.cell(row, col, value)
    path = tmp_path / "synthetic.xlsm"
    wb.save(path)
    wb.close()
    args = dict(patients=patients, sessions=sessions)
    result = interpret_daily_input_grid(grid, **args)
    assert result == read_daily_input(path, **args)
    assert [c.source_row for c in result.source_changes] == [7, 13, 14]
    assert result.absences[0].end_time == time(10)
    assert result.absences[1].absence_kind == AbsenceKind.PATIENT
    assert result.cancellations[0].target_session_id == "S2"


def test_later_date_preserves_exact_date_binding():
    grid, patients, sessions = fixture(NEXT)
    result = interpret_daily_input_grid(grid, patients=patients, sessions=sessions)
    assert result.target_date == NEXT
    assert all(a.absence_date == NEXT for a in result.absences)
    assert all(c.cancellation_date == NEXT for c in result.cancellations)
    assert all(NEXT.isoformat() in c.change_id for c in result.source_changes)


def test_incomplete_draft_is_not_an_applied_operational_change():
    grid, patients, sessions = fixture()
    partial = DailyInputDayDraft(DAY, grid.therapist_rows,
                                 (("DEMO P1", None, None, None, None),))
    with pytest.raises(DailyInputReadError, match="Κατάσταση"):
        interpret_daily_input_grid(partial, patients=patients, sessions=sessions)


def test_clean_and_restore_use_existing_options():
    grid, patients, sessions = fixture()
    result = interpret_daily_input_grid(grid, patients=patients, sessions=sessions,
                                        ignore_rows=(7, 13, 14), allow_empty=True)
    assert not result.absences and not result.cancellations
    with pytest.raises(DailyInputReadError, match="no active"):
        interpret_daily_input_grid(grid.empty_for(DAY), patients=patients, sessions=sessions)
