from datetime import date, datetime

from rehab_excel.daily_input_sheet import _excel_datetime


def test_excel_datetime_converts_date_to_midnight_datetime():
    value = _excel_datetime(date(2026, 9, 19))
    assert value == datetime(2026, 9, 19, 0, 0, 0)
    assert isinstance(value, datetime)
