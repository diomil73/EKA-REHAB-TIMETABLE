from datetime import date

from rehab_excel.student_registration import (
    choose_student_target_row,
    excel_date_serial,
    resolve_boolean_cell_value,
)


def test_student_target_row_starts_at_first_data_row():
    assert choose_student_target_row(1, 1, 1) == 2


def test_student_target_row_follows_last_real_registry_value():
    assert choose_student_target_row(7, 7, 7) == 8


def test_student_target_row_uses_furthest_identity_column():
    assert choose_student_target_row(7, 9, 8) == 10


def test_student_boolean_values_use_configured_yes_no_labels():
    configured = ("ΝΑΙ", "ΟΧΙ")
    assert resolve_boolean_cell_value(True, configured) == "ΝΑΙ"
    assert resolve_boolean_cell_value(False, configured) == "ΟΧΙ"


def test_student_boolean_values_have_safe_fallbacks():
    assert resolve_boolean_cell_value(True, ()) == "ΝΑΙ"
    assert resolve_boolean_cell_value(False, ()) == "ΟΧΙ"


def test_excel_date_serial_preserves_calendar_day_without_datetime_conversion():
    assert excel_date_serial(date(2026, 10, 1)) == 46296
    assert excel_date_serial(date(2026, 12, 31)) == 46387
