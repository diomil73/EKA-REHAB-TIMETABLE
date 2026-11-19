from datetime import date, datetime, time

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

from rehab_excel.daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError
from rehab_excel.daily_input_draft_archive import (
    ARCHIVE_SHEET, HEADERS, archive_rows, capture_daily_input,
    decode_archive_rows, read_archive_sheet,
)

DAY = date(2026, 10, 26)
NEXT = date(2026, 10, 27)


def sheet():
    ws = Workbook().active
    ws.title = "DAILY_INPUT"
    ws["B2"] = "2026-10-26"
    for row, headings in (
        (6, ("Θεραπευτής", "Όλη ημέρα", "Από", "Έως", "Αιτία", "Σχόλιο")),
        (12, ("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο")),
    ):
        for col, heading in enumerate(headings, 1):
            ws.cell(row, col, heading)
    ws["A11"] = "ΑΠΟΥΣΙΕΣ / ΑΚΥΡΩΣΕΙΣ ΑΣΘΕΝΩΝ"
    for name, area in (("PY_THERAPISTS", "A7:A8"), ("PY_PATIENTS", "A13:A15")):
        dv = DataValidation(type="list", formula1="=" + name)
        dv.add(area)
        ws.add_data_validation(dv)
    return ws


def capture(ws):
    return capture_daily_input(ws, parse_date=date.fromisoformat)


def test_capture_dynamic_ranges_including_partial_rows_and_literal_formula_text():
    ws = sheet()
    ws["A7"] = "DEMO T"
    ws["C7"] = time(8, 15)
    ws["D7"] = time(10)
    ws["F8"] = "only a comment"
    ws["A14"] = "DEMO P"
    ws["E14"] = "=literal comment"
    ws["E14"].data_type = "s"
    draft = capture(ws)
    assert draft.service_date == DAY and draft.layout == (2, 3)
    assert draft.therapist_rows == (
        ("DEMO T", None, time(8, 15), time(10), None, None),
        (None, None, None, None, None, "only a comment"),
    )
    assert draft.patient_rows[1] == ("DEMO P", None, None, None, "=literal comment")


@pytest.mark.parametrize("case", ["formula", "error", "header", "tail", "gap", "range"])
def test_capture_rejects_unsafe_layout_or_data(case):
    ws = sheet()
    if case == "formula":
        ws["E14"] = "=1+1"
    elif case == "error":
        ws["C7"] = "#REF!"
    elif case == "header":
        ws["C6"] = "wrong"
    elif case == "tail":
        ws["E16"] = "outside form"
    elif case == "gap":
        ws["A9"] = "outside form"
    else:
        ws.data_validations.dataValidation[0].add("A9")
    with pytest.raises(DailyInputDaySwitchError):
        capture(ws)


def drafts():
    return (
        DailyInputDayDraft(DAY, (("Δοκιμή", False, time(8, 15), 0.5, None, "=text"),),
                           (("DEMO P", True, datetime(2026, 10, 26, 9), date(2026, 10, 26), ""),)),
        DailyInputDayDraft(NEXT, ((None, None, None, None, None, "α" * 20000),),
                           ((None,) * 5,)),
    )


def test_exact_typed_archive_roundtrip_and_multichunk_sheet_readback():
    original = drafts()
    rows = archive_rows(reversed(original))
    assert len(rows) > 2
    assert all(len(row[4]) <= 24004 and row[4].startswith("b64:") for row in rows)
    assert decode_archive_rows(reversed(rows)) == original
    ws = Workbook().active
    ws.title = ARCHIVE_SHEET
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)
    actual = read_archive_sheet(ws)
    assert actual == original
    assert type(actual[0].therapist_rows[0][1]) is bool
    assert type(actual[0].patient_rows[0][2]) is datetime
    assert type(actual[0].patient_rows[0][3]) is date


@pytest.mark.parametrize("case", ["duplicate", "missing", "payload", "date", "count", "digest"])
def test_corrupt_archive_fails_closed(case):
    rows = list(archive_rows(drafts()))
    if case == "duplicate":
        rows.append(rows[0])
    elif case == "missing":
        rows.pop()
    elif case == "payload":
        rows[0] = (*rows[0][:4], "b64:AAAA")
    elif case == "date":
        rows[0] = ("2026-10-28", *rows[0][1:])
    elif case == "count":
        row = rows[0]
        rows[0] = (*row[:2], row[2] + 1, *row[3:])
    else:
        row = rows[0]
        rows[0] = (*row[:3], "0" * 64, row[4])
    with pytest.raises(DailyInputDaySwitchError):
        decode_archive_rows(rows)


def test_duplicate_draft_and_archive_formula_rejected():
    d = drafts()[0]
    with pytest.raises(DailyInputDaySwitchError):
        archive_rows((d, d))
    ws = Workbook().active
    ws.title = ARCHIVE_SHEET
    ws.append(HEADERS)
    ws.append(archive_rows((d,))[0])
    ws["E2"] = "=1+1"
    with pytest.raises(DailyInputDaySwitchError):
        read_archive_sheet(ws)

def test_archive_survives_synthetic_xlsx_serialization():
    from io import BytesIO
    from openpyxl import load_workbook

    original = drafts()
    wb = Workbook()
    ws = wb.create_sheet(ARCHIVE_SHEET)
    ws.sheet_state = "veryHidden"
    ws.protection.sheet = True
    ws.append(HEADERS)
    for row in archive_rows(original):
        ws.append(row)
    buffer = BytesIO()
    wb.save(buffer)
    wb.close()
    buffer.seek(0)
    reopened = load_workbook(buffer, data_only=False)
    try:
        archive = reopened[ARCHIVE_SHEET]
        assert archive.sheet_state == "veryHidden"
        assert archive.protection.sheet
        assert archive_rows(read_archive_sheet(archive)) == archive_rows(original)
    finally:
        reopened.close()


def test_excessive_declared_validation_range_rejected_before_grid_capture():
    ws = sheet()
    ws.data_validations.dataValidation[1].sqref = "A13:A1048576"
    with pytest.raises(DailyInputDaySwitchError, match="explicit layout review"):
        capture(ws)
