"""Read-only integration tests using a synthetic, VBA-free OOXML fixture.

The XLSM extension exercises the application's input contract, not native Excel.
"""
from datetime import date
from hashlib import sha256

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

from rehab_excel.daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError
from rehab_excel.daily_input_draft_archive import ARCHIVE_SHEET, HEADERS, archive_rows
from rehab_excel.daily_input_switch_preparation import prepare_open_day_switch
from rehab_excel.finalized_day_snapshots import HEADERS as HISTORY_HEADERS

DAY = date(2026, 10, 26)
NEXT = date(2026, 10, 27)


def source(tmp_path, *, archive=True, protected=True, finalized=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "DAILY_INPUT"
    ws["B2"] = "26/10/2026"
    for row, headings in (
        (6, ("Θεραπευτής", "Όλη ημέρα", "Από", "Έως", "Αιτία", "Σχόλιο")),
        (11, ("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο")),
    ):
        for col, heading in enumerate(headings, 1):
            ws.cell(row, col, heading)
    ws["A10"] = "ΑΠΟΥΣΙΕΣ / ΑΚΥΡΩΣΕΙΣ ΑΣΘΕΝΩΝ"
    for name, area in (("PY_THERAPISTS", "A7"), ("PY_PATIENTS", "A12")):
        dv = DataValidation(type="list", formula1="=" + name)
        dv.add(area)
        ws.add_data_validation(dv)
    ws["A7"] = "DEMO T1"
    ws["F7"] = "partial"
    saved = DailyInputDayDraft(NEXT, (("DEMO T2", "ΝΑΙ", None, None, None, None),),
                              ((None,) * 5,))
    if archive:
        store = wb.create_sheet(ARCHIVE_SHEET)
        store.sheet_state = "veryHidden"
        store.protection.sheet = protected
        store.append(HEADERS)
        for row in archive_rows((saved,)):
            store.append(row)
    if finalized:
        history = wb.create_sheet("FINALIZED_DAY_SNAPSHOTS")
        history.append(HISTORY_HEADERS)
        history.append(("synthetic-final", NEXT.isoformat(), 1, "CURRENT_FINAL",
                        "2026-10-27T16:00:00+02:00", None, 0, 0, 0, 0, 0, 0, 0, 0, 0))
    path = tmp_path / "synthetic.xlsm"
    wb.save(path)
    wb.close()
    return path, saved


def test_preparation_binds_saved_archive_and_never_changes_source(tmp_path):
    path, saved = source(tmp_path)
    before = sha256(path.read_bytes()).hexdigest()
    plan = prepare_open_day_switch(path, expected_current_date=DAY, target_date=NEXT)
    assert plan.source_sha256 == before
    assert plan.day_plan.incoming == saved
    assert plan.day_plan.outgoing.therapist_rows[0][-1] == "partial"
    assert plan.expected_archive_rows == archive_rows(plan.day_plan.saved_after)
    assert sha256(path.read_bytes()).hexdigest() == before


@pytest.mark.parametrize("case", ["missing", "unprotected", "finalized", "reinitialize", "stale"])
def test_preparation_stops_unsafe_transition_without_writes(tmp_path, case):
    path, _ = source(tmp_path, archive=case != "missing",
                     protected=case != "unprotected", finalized=case == "finalized")
    before = path.read_bytes()
    with pytest.raises(DailyInputDaySwitchError):
        prepare_open_day_switch(
            path, expected_current_date=NEXT if case == "stale" else DAY,
            target_date=NEXT, initialize_archive=case == "reinitialize",
        )
    assert path.read_bytes() == before


def test_archive_initialization_requires_explicit_first_install(tmp_path):
    path, _ = source(tmp_path, archive=False)
    before = path.read_bytes()
    plan = prepare_open_day_switch(path, expected_current_date=DAY, target_date=NEXT,
                                   initialize_archive=True)
    assert plan.day_plan.saved_before == ()
    assert plan.day_plan.incoming == plan.day_plan.outgoing.empty_for(NEXT)
    assert path.read_bytes() == before
