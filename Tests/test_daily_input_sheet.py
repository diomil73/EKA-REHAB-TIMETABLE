from datetime import date, time

from rehab_excel.daily_input_sheet import build_daily_input_spec


def test_daily_input_spec_keeps_patient_status_options_exactly():
    statuses = ("ΠΑΡΩΝ", "εκτός κλινικής", "Αναβολή ΦΘ")
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("Α", "Β"),
        patient_names=("ΑΣΘΕΝΗΣ 1", "ΑΣΘΕΝΗΣ 2"),
        patient_statuses=statuses,
        timeslots=(time(8, 30), time(9, 15)),
    )
    assert spec.patient_statuses == statuses


def test_daily_input_spec_deduplicates_dropdown_values_without_reordering():
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("Α", "Α", "Β"),
        patient_names=("Π1", "Π2", "Π1"),
        patient_statuses=("ΠΑΡΩΝ", "ΠΑΡΩΝ", "ΑΠΩΝ"),
        timeslots=(time(8, 30), time(9, 15)),
    )
    assert spec.therapist_names == ("Α", "Β")
    assert spec.patient_names == ("Π1", "Π2")
    assert spec.patient_statuses == ("ΠΑΡΩΝ", "ΑΠΩΝ")


def test_daily_input_layout_has_separate_therapist_and_patient_sections():
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("Α",),
        patient_names=("Π1",),
        patient_statuses=("ΠΑΡΩΝ",),
        timeslots=(time(8, 30),),
        therapist_rows=10,
        patient_rows=20,
    )
    assert spec.therapist_first_row == 7
    assert spec.therapist_last_row == 16
    assert spec.patient_header_row == 20
    assert spec.patient_first_row == 21
    assert spec.patient_last_row == 40
