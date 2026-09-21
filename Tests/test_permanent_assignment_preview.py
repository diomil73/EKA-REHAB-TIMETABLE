from datetime import time

from rehab_excel.permanent_preview import excel_column_name, parse_hhmm


def test_parse_hhmm():
    assert parse_hhmm("13:00") == time(13, 0)


def test_excel_column_name():
    assert excel_column_name(1) == "A"
    assert excel_column_name(8) == "H"
    assert excel_column_name(26) == "Z"
    assert excel_column_name(27) == "AA"
