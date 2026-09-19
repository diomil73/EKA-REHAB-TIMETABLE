from datetime import date, time

from rehab_excel.daily_input_sheet import build_daily_input_spec


def test_timeslots_are_rendered_as_hhmm_text():
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("A",),
        patient_names=("P",),
        patient_statuses=("ΠΑΡΩΝ",),
        timeslots=(time(8, 30), time(9, 15), time(13, 0)),
    )
    assert spec.timeslots == ("08:30", "09:15", "13:00")
